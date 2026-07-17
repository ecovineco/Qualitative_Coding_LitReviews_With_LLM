"""LLM API client module with a provider-agnostic interface.

This module defines an abstract base class (``BaseLLMClient``) that any LLM
provider must implement, plus a concrete ``AnthropicClient`` for Claude.

To add a new provider (e.g. OpenAI, Google Gemini):
    1. Create a new class that inherits from ``BaseLLMClient``.
    2. Implement all abstract methods.
    3. Register it in ``get_llm_client()``.

Typical usage::

    from api import get_llm_client
    client = get_llm_client("anthropic")
    file_id = client.upload_file("paper.pdf")

Architecture:
    The ``BaseLLMClient`` defines four operations that the pipeline needs:
        - upload_file   → Send a PDF, get back an identifier.
        - submit_batch  → Send multiple analysis requests at once.
        - check_batch   → Poll for batch completion.
        - get_results   → Download and return raw results.

    Each provider implements these differently, but the pipeline (main.py)
    only ever interacts with the abstract interface.
"""

from __future__ import annotations

import abc
import json
import logging
import time
from collections import deque
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import re
import requests

import config

logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════
# Shared helpers
# ══════════════════════════════════════════════════════════════════════


def _make_custom_id(filename: str, idx: int) -> str:
    """Build the batch custom_id for a given file.

    This function is the single source of truth for custom_id generation.
    It is used both when submitting the batch (so the API sees this id)
    and when rebuilding the ``custom_id -> filename`` mapping after
    results come back.  The two sites MUST agree, so they both call this.

    Args:
        filename: The original PDF filename.
        idx: 1-based index of the file within the batch.

    Returns:
        A sanitised custom_id string, truncated to 64 characters
        (the Anthropic Batch API limit).
    """
    safe_name = re.sub(r"[^a-zA-Z0-9_-]", "_", filename)
    return f"req-{idx:04d}-{safe_name}"[:64]


def _retry_after_seconds(exc: Exception) -> Optional[float]:
    """Best-effort extraction of a ``Retry-After`` value (seconds) from an error.

    Azure/OpenAI include a ``Retry-After`` header on HTTP 429 responses
    telling the caller exactly how long to wait. Used to back off
    precisely when the pre-emptive rate limiter still gets overruled.

    Args:
        exc: The exception raised for the failed request (expected to be
            an ``openai.RateLimitError`` or similar with a ``.response``).

    Returns:
        The header value in seconds, or ``None`` if absent/unparsable —
        callers should fall back to a sensible default in that case.
    """
    response = getattr(exc, "response", None)
    if response is None:
        return None
    header = getattr(response, "headers", {}).get("retry-after")
    if not header:
        return None
    try:
        return float(header)
    except (TypeError, ValueError):
        return None


# ══════════════════════════════════════════════════════════════════════
# Data structures
# ══════════════════════════════════════════════════════════════════════


@dataclass
class BatchStatus:
    """Represents the current state of a batch job.

    Attributes:
        batch_id: The unique identifier for the batch.
        is_complete: ``True`` if the batch has finished processing.
        results_url: URL to download results (only set when complete).
        counts: Dictionary with request counts by status
            (e.g. ``{"succeeded": 10, "errored": 1}``).
        raw: The full raw response from the API for debugging.
    """

    batch_id: str
    is_complete: bool
    results_url: Optional[str] = None
    counts: Optional[Dict[str, int]] = None
    raw: Optional[Dict[str, Any]] = None


@dataclass
class BatchResultItem:
    """A single result from a completed batch.

    Attributes:
        custom_id: The identifier you assigned to this request.
        success: ``True`` if the request succeeded.
        text: The LLM's text response (``None`` if failed).
        error_message: Description of the error (``None`` if succeeded).
        raw: The full raw result item for debugging.
    """

    custom_id: str
    success: bool
    text: Optional[str] = None
    error_message: Optional[str] = None
    raw: Optional[Dict[str, Any]] = None


# ══════════════════════════════════════════════════════════════════════
# Abstract base class
# ══════════════════════════════════════════════════════════════════════


class BaseLLMClient(abc.ABC):
    """Abstract interface for LLM provider clients.

    Any LLM provider used by this pipeline must implement this interface.
    The pipeline (``main.py``) only calls these methods, so swapping
    providers requires zero changes to the rest of the codebase.
    """

    @abc.abstractmethod
    def upload_file(self, file_path: str | Path) -> str:
        """Upload a PDF file and return a provider-specific file identifier.

        Args:
            file_path: Path to the PDF file on disk.

        Returns:
            A string identifier that can be used to reference this file
            in subsequent API calls.

        Raises:
            RuntimeError: If the upload fails after all retry attempts.
        """

    @abc.abstractmethod
    def submit_batch(
        self,
        file_rows: List[Tuple[str, str]],
        prompt: str,
    ) -> str:
        """Submit a batch of analysis requests.

        Each request pairs the given prompt with one uploaded document.

        Args:
            file_rows: List of ``(filename, file_id)`` tuples.
            prompt: The analysis prompt text to send with each document.

        Returns:
            A batch identifier string for tracking the job.

        Raises:
            RuntimeError: If the batch submission fails.
        """

    @abc.abstractmethod
    def check_batch(self, batch_id: str) -> BatchStatus:
        """Check the current status of a batch job.

        Args:
            batch_id: The identifier returned by ``submit_batch()``.

        Returns:
            A ``BatchStatus`` object describing the current state.

        Raises:
            RuntimeError: If the status check fails.
        """

    @abc.abstractmethod
    def get_results(self, batch_status: BatchStatus) -> List[BatchResultItem]:
        """Download and parse results from a completed batch.

        Args:
            batch_status: A ``BatchStatus`` with ``is_complete=True``.

        Returns:
            A list of ``BatchResultItem`` objects, one per request.

        Raises:
            RuntimeError: If the download or parsing fails.
            ValueError: If the batch is not yet complete.
        """

    def poll_until_complete(
        self,
        batch_id: str,
        poll_interval: int = config.POLL_INTERVAL_SECONDS,
    ) -> BatchStatus:
        """Block until a batch job finishes, polling at regular intervals.

        This is a convenience method built on ``check_batch()``.

        Args:
            batch_id: The identifier returned by ``submit_batch()``.
            poll_interval: Seconds between status checks.

        Returns:
            The final ``BatchStatus`` with ``is_complete=True``.
        """
        logger.info("Polling batch '%s' every %ds...", batch_id, poll_interval)
        while True:
            status = self.check_batch(batch_id)
            logger.info(
                "Batch '%s' — complete=%s, counts=%s",
                batch_id,
                status.is_complete,
                status.counts,
            )
            if status.is_complete:
                return status
            time.sleep(poll_interval)


# ══════════════════════════════════════════════════════════════════════
# Anthropic implementation
# ══════════════════════════════════════════════════════════════════════


class AnthropicClient(BaseLLMClient):
    """Concrete LLM client for Anthropic's Claude API.

    Implements file upload via the Files API (beta), batch processing via
    the Message Batches API, and JSONL result retrieval.

    Attributes:
        api_key: The Anthropic API key.
        model: The model identifier to use for analysis.
        max_tokens: Maximum tokens per response.
        base_url: Base URL for all API endpoints.
    """

    def __init__(
        self,
        api_key: str = "",
        model: str = config.ANTHROPIC_MODEL,
        max_tokens: int = config.ANTHROPIC_MAX_TOKENS,
        base_url: str = config.ANTHROPIC_BASE_URL,
    ) -> None:
        """Initialise the Anthropic client.

        Args:
            api_key: Anthropic API key.
            model: Model identifier (e.g. "claude-sonnet-4-20250514").
            max_tokens: Max tokens per response.
            base_url: API base URL.

        Raises:
            ValueError: If the API key is empty.
        """
        if not api_key:
            api_key = config.ANTHROPIC_API_KEY
        if not api_key:
            raise ValueError(
                "Anthropic API key is missing. Set the ANTHROPIC_API_KEY "
                "environment variable or pass it directly."
            )
        self.api_key = api_key
        self.model = model
        self.max_tokens = max_tokens
        self.base_url = base_url.rstrip("/")

    # ── Internal helpers ────────────────────────────────────────────

    def _headers(self, include_files_beta: bool = False) -> Dict[str, str]:
        """Build HTTP headers for Anthropic API requests.

        Args:
            include_files_beta: If ``True``, include the beta header
                required by the Files API.

        Returns:
            Dictionary of HTTP headers.
        """
        headers = {
            "x-api-key": self.api_key,
            "anthropic-version": config.ANTHROPIC_VERSION,
        }
        if include_files_beta:
            headers["anthropic-beta"] = config.ANTHROPIC_FILES_BETA
        return headers

    def _request(
        self,
        method: str,
        endpoint: str,
        *,
        headers: Optional[Dict[str, str]] = None,
        json_body: Optional[Any] = None,
        files: Optional[Any] = None,
        timeout: int = config.REQUEST_TIMEOUT,
    ) -> Dict[str, Any]:
        """Send an HTTP request to the Anthropic API.

        This is the single point of contact for all HTTP calls, making it
        easy to add logging, retry logic, or rate-limit handling.

        Args:
            method: HTTP method ("GET" or "POST").
            endpoint: API endpoint path (e.g. "/v1/files").
            headers: HTTP headers to send.
            json_body: JSON-serialisable request body (for POST).
            files: Multipart file data (for POST with file upload).
            timeout: Request timeout in seconds.

        Returns:
            Parsed JSON response as a dictionary.

        Raises:
            RuntimeError: If the response status code indicates an error.
        """
        url = f"{self.base_url}{endpoint}"
        logger.debug("%s %s", method, url)

        resp = requests.request(
            method=method,
            url=url,
            headers=headers,
            json=json_body,
            files=files,
            timeout=timeout,
        )

        if resp.status_code >= 300:
            raise RuntimeError(
                f"API error: {method} {endpoint} returned "
                f"{resp.status_code}:\n{resp.text}"
            )

        return resp.json()

    def _download_raw(
        self,
        url: str,
        timeout: int = 600,
    ) -> str:
        """Download raw text content from a URL (authenticated).

        Used to fetch JSONL result files from Anthropic's servers.

        Args:
            url: The full URL to download.
            timeout: Request timeout in seconds.

        Returns:
            The raw response text.

        Raises:
            RuntimeError: If the download fails.
        """
        headers = self._headers(include_files_beta=True)
        resp = requests.get(url, headers=headers, timeout=timeout)

        if resp.status_code >= 300:
            raise RuntimeError(
                f"Download failed: {resp.status_code}:\n{resp.text}"
            )

        return resp.text

    # ── Public interface (BaseLLMClient) ────────────────────────────

    def upload_file(self, file_path: str | Path) -> str:
        """Upload a PDF to Anthropic's Files API.

        Retries on failure according to ``config.UPLOAD_RETRIES``.

        Args:
            file_path: Path to the PDF file.

        Returns:
            The Anthropic file ID string.

        Raises:
            RuntimeError: If upload fails after all retries.
        """
        path = Path(file_path)
        last_error: Optional[Exception] = None

        for attempt in range(1, config.UPLOAD_RETRIES + 1):
            try:
                with open(path, "rb") as f:
                    multipart_files = {
                        "file": (path.name, f, "application/pdf"),
                    }
                    data = self._request(
                        "POST",
                        "/v1/files",
                        headers=self._headers(include_files_beta=True),
                        files=multipart_files,
                    )

                file_id = data.get("id", "")
                if not file_id:
                    raise RuntimeError(
                        f"No file_id in response for '{path.name}': {data}"
                    )

                logger.info(
                    "Uploaded '%s' -> %s (attempt %d)",
                    path.name,
                    file_id,
                    attempt,
                )
                return file_id

            except Exception as exc:
                last_error = exc
                logger.warning(
                    "Upload attempt %d/%d failed for '%s': %s",
                    attempt,
                    config.UPLOAD_RETRIES,
                    path.name,
                    exc,
                )
                if attempt < config.UPLOAD_RETRIES:
                    time.sleep(config.UPLOAD_RETRY_DELAY)

        raise RuntimeError(
            f"Upload failed for '{path.name}' after {config.UPLOAD_RETRIES} "
            f"attempts. Last error: {last_error}"
        )

    def submit_batch(
        self,
        file_rows: List[Tuple[str, str]],
        prompt: str,
    ) -> str:
        """Submit a Message Batch to Anthropic.

        Constructs one request per file, each containing the analysis
        prompt and a document reference.

        Args:
            file_rows: List of ``(filename, file_id)`` tuples.
            prompt: The analysis prompt text.

        Returns:
            The batch ID string.

        Raises:
            RuntimeError: If batch creation fails.
        """
        batch_requests: List[Dict[str, Any]] = []

        for idx, (filename, file_id) in enumerate(file_rows, start=1):
            custom_id = _make_custom_id(filename, idx)

            request_item = {
                "custom_id": custom_id,
                "params": {
                    "model": self.model,
                    "max_tokens": self.max_tokens,
                    "messages": [
                        {
                            "role": "user",
                            "content": [
                                {"type": "text", "text": prompt},
                                {
                                    "type": "document",
                                    "source": {
                                        "type": "file",
                                        "file_id": file_id,
                                    },
                                },
                            ],
                        }
                    ],
                },
            }
            batch_requests.append(request_item)

        logger.info("Submitting batch with %d requests...", len(batch_requests))

        headers = {
            **self._headers(include_files_beta=True),
            "Content-Type": "application/json",
        }
        data = self._request(
            "POST",
            "/v1/messages/batches",
            headers=headers,
            json_body={"requests": batch_requests},
        )

        batch_id = data.get("id", "")
        if not batch_id:
            raise RuntimeError(f"No batch_id in response: {data}")

        logger.info("Batch submitted: %s", batch_id)
        return batch_id

    def check_batch(self, batch_id: str) -> BatchStatus:
        """Check the status of an Anthropic Message Batch.

        Args:
            batch_id: The batch identifier.

        Returns:
            A ``BatchStatus`` object.
        """
        data = self._request(
            "GET",
            f"/v1/messages/batches/{batch_id}",
            headers=self._headers(include_files_beta=True),
        )

        processing_status = data.get("processing_status", "")
        is_complete = processing_status == "ended"

        return BatchStatus(
            batch_id=batch_id,
            is_complete=is_complete,
            results_url=data.get("results_url"),
            counts=data.get("request_counts"),
            raw=data,
        )

    def get_results(self, batch_status: BatchStatus) -> List[BatchResultItem]:
        """Download and parse results from a completed Anthropic batch.

        Args:
            batch_status: A completed ``BatchStatus``.

        Returns:
            List of ``BatchResultItem`` objects.

        Raises:
            ValueError: If the batch is not complete or has no results URL.
        """
        if not batch_status.is_complete:
            raise ValueError(
                f"Batch '{batch_status.batch_id}' is not yet complete."
            )
        if not batch_status.results_url:
            raise ValueError(
                f"Batch '{batch_status.batch_id}' has no results_url."
            )

        logger.info("Downloading results from %s", batch_status.results_url)
        raw_text = self._download_raw(batch_status.results_url)

        items: List[BatchResultItem] = []
        for line in raw_text.strip().splitlines():
            if not line.strip():
                continue

            record = json.loads(line)
            custom_id = record.get("custom_id", "")
            result = record.get("result", {})
            result_type = result.get("type", "")

            if result_type == "succeeded":
                # Extract text from content blocks.
                message = result.get("message", {})
                blocks = message.get("content", [])
                text_parts = [
                    b["text"]
                    for b in blocks
                    if isinstance(b, dict) and b.get("type") == "text"
                ]
                text = "\n".join(text_parts).strip() or None

                items.append(
                    BatchResultItem(
                        custom_id=custom_id,
                        success=True,
                        text=text,
                        raw=record,
                    )
                )
            else:
                error_obj = result.get("error", {})
                error_msg = (
                    error_obj.get("message", "Unknown error")
                    if isinstance(error_obj, dict)
                    else str(error_obj)
                )
                items.append(
                    BatchResultItem(
                        custom_id=custom_id,
                        success=False,
                        error_message=error_msg,
                        raw=record,
                    )
                )

        logger.info(
            "Parsed %d results (%d succeeded, %d failed)",
            len(items),
            sum(1 for i in items if i.success),
            sum(1 for i in items if not i.success),
        )
        return items


# ══════════════════════════════════════════════════════════════════════
# Azure OpenAI implementation
# ══════════════════════════════════════════════════════════════════════


class _AzureTokenRateLimiter:
    """Paces requests to stay within a tokens-per-minute (TPM) quota.

    Azure enforces its quota over a rolling window, so this mirrors that:
    it remembers how many tokens each recent request *actually* used
    (from the API's own ``usage`` field) and, before every new request,
    checks whether sending it would push the trailing-60-second total
    over budget. If so, it sleeps just long enough for the oldest usage
    to age out of the window, then re-checks — i.e. it only delays when
    a request would breach the limit, and lets requests through right
    away whenever there's headroom.

    The size of an upcoming request isn't known ahead of time (it
    depends on the PDF being sent), so the wait check estimates it from
    a running average of the last few *actual* request sizes. The very
    first request of a run always fires immediately, since there's no
    usage history yet to judge it against.

    This works the same way regardless of how small or large the quota
    is: even when a single request's real usage exceeds the entire
    per-minute budget (common with a small quota and large PDFs), the
    limiter simply waits out however much of the window is needed after
    that request before allowing the next one.
    """

    def __init__(
        self,
        tokens_per_minute: int,
        safety_margin: float = 1.0,
        window_seconds: float = 60.0,
        history_size: int = 5,
    ) -> None:
        """Set up the limiter.

        Args:
            tokens_per_minute: The quota to stay under (read from
                config; never hard-coded here).
            safety_margin: Fraction (0-1] of the quota to actually treat
                as the budget, leaving headroom for estimation error.
            window_seconds: Length of the rolling window, in seconds.
            history_size: How many recent request sizes to average when
                estimating the next one.
        """
        self.window_seconds = window_seconds
        self.budget = max(int(tokens_per_minute * safety_margin), 1)
        self._usage: deque = deque()  # (monotonic_timestamp, tokens) still inside the window
        self._recent: deque = deque(maxlen=history_size)  # actual sizes of the last few requests

    def _window_total(self, now: float) -> int:
        """Drop expired usage entries and return the remaining total."""
        cutoff = now - self.window_seconds
        while self._usage and self._usage[0][0] < cutoff:
            self._usage.popleft()
        return sum(tokens for _, tokens in self._usage)

    def throttle(self) -> None:
        """Block, if necessary, until there's room for another request."""
        estimated = int(sum(self._recent) / len(self._recent)) if self._recent else 0
        # A request that actually lands never costs zero tokens, so floor
        # the estimate at 1 — keeps the boundary check below meaningful
        # even in edge cases where the rolling average would round to 0.
        estimated = max(estimated, 1)

        while True:
            now = time.monotonic()
            used = self._window_total(now)

            if not self._usage or used + estimated <= self.budget:
                return  # No usage history yet, or plenty of headroom — go now.

            wait_for = max((self._usage[0][0] + self.window_seconds) - now, 1.0)
            logger.info(
                "Azure TPM pacing: ~%d/%d tokens used in the trailing %ds "
                "(next request estimated at %d tokens) — waiting %.1fs "
                "for the quota to free up...",
                used,
                self.budget,
                int(self.window_seconds),
                estimated,
                wait_for,
            )
            time.sleep(wait_for)

    def record_usage(self, tokens: int) -> None:
        """Log the tokens an actually-completed request used."""
        if tokens <= 0:
            return
        self._usage.append((time.monotonic(), tokens))
        self._recent.append(tokens)


class AzureOpenAIClient(BaseLLMClient):
    """Concrete LLM client for Azure OpenAI (GPT family).

    Azure OpenAI does not expose a batch-discount endpoint with the same
    asynchronous, server-side semantics as Anthropic.  To preserve the
    ``BaseLLMClient`` interface, this client executes each request
    synchronously inside ``submit_batch()`` and persists the results to
    disk as JSONL at ``OUTPUT_DIR/azure_batches/{batch_id}.jsonl``.

    ``check_batch()`` and ``get_results()`` then read from that file, so
    ``--resume`` works exactly the same way as with Anthropic: any
    category whose ``batch_id`` was already persisted is loaded back
    without re-running the LLM.

    PDFs are passed inline as base64-encoded ``input_file`` content
    blocks via the Responses API (``client.responses.create``), which is
    supported on GPT-5 family Azure deployments.  No separate "upload"
    step is needed; ``upload_file()`` simply verifies that the PDF
    exists on disk and returns its absolute path, which is then used as
    the ``file_id`` for that document.

    Requests are paced by an internal sliding-window rate limiter
    (``_AzureTokenRateLimiter``) so that total token usage stays within
    ``config.AZURE_OPENAI_TPM_LIMIT`` tokens per rolling 60-second
    window, waiting only when a request would otherwise exceed it. If
    Azure still returns a 429, the call is retried with backoff up to
    ``config.AZURE_RATE_LIMIT_MAX_RETRIES`` times.

    Attributes:
        api_key: The Azure OpenAI API key.
        endpoint: The Azure OpenAI resource endpoint.
        api_version: API version string.
        model: Deployment name to use for analysis.
        rate_limiter: Paces requests to the configured TPM quota.
    """

    def __init__(
        self,
        api_key: str = "",
        endpoint: str = config.AZURE_OPENAI_ENDPOINT,
        api_version: str = config.AZURE_OPENAI_API_VERSION,
        model: str = config.AZURE_OPENAI_MODEL,
    ) -> None:
        """Initialise the Azure OpenAI client.

        Raises:
            ValueError: If the API key is missing.
            RuntimeError: If the ``openai`` package is not installed.
        """
        if not api_key:
            api_key = config.AZURE_OPENAI_API_KEY
        if not api_key:
            raise ValueError(
                "Azure OpenAI API key is missing. Set the "
                "AZURE_OPENAI_API_KEY environment variable or pass it "
                "directly."
            )
        try:
            from openai import AzureOpenAI  # noqa: F401
        except ImportError as exc:
            raise RuntimeError(
                "The 'openai' package is required for the Azure provider. "
                "Install it with: pip install openai"
            ) from exc

        from openai import AzureOpenAI

        self.api_key = api_key
        self.endpoint = endpoint
        self.api_version = api_version
        self.model = model
        self._client = AzureOpenAI(
            api_key=api_key,
            azure_endpoint=endpoint,
            api_version=api_version,
        )
        self.rate_limiter = _AzureTokenRateLimiter(
            tokens_per_minute=config.AZURE_OPENAI_TPM_LIMIT,
            safety_margin=config.AZURE_RATE_LIMIT_SAFETY_MARGIN,
        )

    # ── Internal helpers ────────────────────────────────────────────

    def _batches_dir(self) -> Path:
        """Return the directory where Azure batch JSONL files live."""
        path = Path(config.OUTPUT_DIR) / "azure_batches"
        path.mkdir(parents=True, exist_ok=True)
        return path

    def _batch_file(self, batch_id: str) -> Path:
        """Return the JSONL file path for a given batch_id."""
        return self._batches_dir() / f"{batch_id}.jsonl"

    @staticmethod
    def _encode_pdf(path: Path) -> str:
        """Read a PDF and return a base64 data URI string."""
        import base64

        with open(path, "rb") as f:
            raw = f.read()
        b64 = base64.b64encode(raw).decode("ascii")
        return f"data:application/pdf;base64,{b64}"

    def _call_with_rate_limit(
        self,
        prompt: str,
        pdf_data_uri: str,
        filename: str,
    ) -> Any:
        """Call the Responses API while respecting the TPM quota.

        Before every attempt, blocks on ``self.rate_limiter`` until the
        trailing 60-second token usage has room for another request.
        After a successful call, records the *actual* tokens the API
        reports so later waits are based on real usage rather than a
        guess.

        If Azure still returns a 429 despite the pre-emptive wait (e.g.
        the estimate for this particular request was too low), backs
        off using the ``Retry-After`` header when present and retries,
        up to ``config.AZURE_RATE_LIMIT_MAX_RETRIES`` attempts.

        Args:
            prompt: The analysis prompt text.
            pdf_data_uri: The base64 ``data:application/pdf`` URI.
            filename: Original filename, passed through to the
                ``input_file`` block and used in retry log messages.

        Returns:
            The raw Responses API result object.

        Raises:
            RuntimeError: If every retry attempt is exhausted.
        """
        from openai import RateLimitError

        last_error: Optional[Exception] = None

        for attempt in range(1, config.AZURE_RATE_LIMIT_MAX_RETRIES + 1):
            self.rate_limiter.throttle()
            try:
                response = self._client.responses.create(
                    model=self.model,
                    input=[
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": "input_text",
                                    "text": prompt,
                                },
                                {
                                    "type": "input_file",
                                    "file_data": pdf_data_uri,
                                    "filename": filename,
                                },
                            ],
                        }
                    ],
                )
                usage = getattr(response, "usage", None)
                self.rate_limiter.record_usage(getattr(usage, "total_tokens", 0) or 0)
                return response

            except RateLimitError as exc:
                last_error = exc
                wait_s = _retry_after_seconds(exc) or self.rate_limiter.window_seconds
                logger.warning(
                    "  Azure TPM limit hit on '%s' (attempt %d/%d) — "
                    "waiting %.1fs before retrying...",
                    filename,
                    attempt,
                    config.AZURE_RATE_LIMIT_MAX_RETRIES,
                    wait_s,
                )
                time.sleep(wait_s)

        raise RuntimeError(
            f"Azure rate limit still exceeded after "
            f"{config.AZURE_RATE_LIMIT_MAX_RETRIES} attempts for "
            f"'{filename}'. Last error: {last_error}"
        )

    # ── Public interface (BaseLLMClient) ────────────────────────────

    def upload_file(self, file_path: str | Path) -> str:
        """Verify a PDF exists and return its absolute path as the file_id.

        Azure OpenAI does not require a separate persistent upload step
        for chat completions; the PDF is attached inline at request time.
        Using the absolute path as the file_id makes the value stable
        across ``--resume`` invocations (the mapping in
        ``uploaded_file_ids.xlsx`` stays valid).

        Args:
            file_path: Path to the PDF file.

        Returns:
            The absolute path to the file as a string.

        Raises:
            FileNotFoundError: If the file does not exist.
        """
        path = Path(file_path).resolve()
        if not path.exists():
            raise FileNotFoundError(f"PDF not found: {path}")
        logger.info("Registered '%s' (Azure: no upload needed)", path.name)
        return str(path)

    def submit_batch(
        self,
        file_rows: List[Tuple[str, str]],
        prompt: str,
    ) -> str:
        """Run a synchronous "batch": one chat-completion call per file.

        Results are written to ``OUTPUT_DIR/azure_batches/{batch_id}.jsonl``
        so that ``check_batch()`` and ``get_results()`` can read them back
        on a subsequent ``--resume``.

        Each call is paced by ``self.rate_limiter`` (see
        ``_AzureTokenRateLimiter``) so the trailing-60-second token usage
        stays within ``config.AZURE_OPENAI_TPM_LIMIT`` — requests fire
        immediately while there's headroom and only wait when they'd
        otherwise breach the quota.

        Args:
            file_rows: List of ``(filename, file_id)`` tuples, where each
                ``file_id`` is the absolute path returned by
                ``upload_file()``.
            prompt: The analysis prompt text.

        Returns:
            A batch identifier string of the form ``azurebatch_<uuid>``.

        Raises:
            RuntimeError: If creating the batch file fails.
        """
        import uuid

        batch_id = f"azurebatch_{uuid.uuid4().hex[:24]}"
        batch_path = self._batch_file(batch_id)
        logger.info(
            "Submitting Azure batch '%s' with %d requests "
            "(synchronous, one call per file)...",
            batch_id,
            len(file_rows),
        )

        with open(batch_path, "w", encoding="utf-8") as out:
            for idx, (filename, file_id) in enumerate(file_rows, start=1):
                custom_id = _make_custom_id(filename, idx)
                record: Dict[str, Any] = {"custom_id": custom_id}

                try:
                    pdf_data_uri = self._encode_pdf(Path(file_id))
                    response = self._call_with_rate_limit(
                        prompt=prompt,
                        pdf_data_uri=pdf_data_uri,
                        filename=filename,
                    )
                    text = response.output_text or ""
                    record["result"] = {
                        "type": "succeeded",
                        "text": text,
                    }
                    logger.info(
                        "  [%d/%d] %s -> %d chars",
                        idx,
                        len(file_rows),
                        filename,
                        len(text),
                    )
                except Exception as exc:
                    record["result"] = {
                        "type": "errored",
                        "error_message": f"{type(exc).__name__}: {exc}",
                    }
                    logger.warning(
                        "  [%d/%d] %s -> FAILED: %s",
                        idx,
                        len(file_rows),
                        filename,
                        exc,
                    )

                out.write(json.dumps(record, ensure_ascii=False) + "\n")
                out.flush()  # Persist incrementally so a crash mid-batch
                             # still leaves partial results on disk.

        logger.info("Azure batch '%s' written to %s", batch_id, batch_path)
        return batch_id

    def check_batch(self, batch_id: str) -> BatchStatus:
        """Return the status of an Azure batch.

        Since Azure batches are processed synchronously inside
        ``submit_batch()``, completion is determined by whether the JSONL
        file exists on disk.

        Args:
            batch_id: The Azure batch identifier.

        Returns:
            A ``BatchStatus`` object.  ``is_complete`` is ``True`` iff
            the batch's JSONL file is present.
        """
        batch_path = self._batch_file(batch_id)
        if not batch_path.exists():
            return BatchStatus(
                batch_id=batch_id,
                is_complete=False,
                counts={"processing": 1},
            )

        succeeded = 0
        errored = 0
        with open(batch_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                rec = json.loads(line)
                rtype = rec.get("result", {}).get("type", "")
                if rtype == "succeeded":
                    succeeded += 1
                else:
                    errored += 1

        return BatchStatus(
            batch_id=batch_id,
            is_complete=True,
            results_url=str(batch_path),
            counts={"succeeded": succeeded, "errored": errored},
        )

    def get_results(self, batch_status: BatchStatus) -> List[BatchResultItem]:
        """Read results from a completed Azure batch JSONL file.

        Args:
            batch_status: A completed ``BatchStatus``.

        Returns:
            List of ``BatchResultItem`` objects, one per request.

        Raises:
            ValueError: If the batch is not complete.
        """
        if not batch_status.is_complete:
            raise ValueError(
                f"Azure batch '{batch_status.batch_id}' is not complete."
            )

        batch_path = self._batch_file(batch_status.batch_id)
        if not batch_path.exists():
            raise RuntimeError(
                f"Azure batch results file missing: {batch_path}"
            )

        items: List[BatchResultItem] = []
        with open(batch_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                rec = json.loads(line)
                custom_id = rec.get("custom_id", "")
                result = rec.get("result", {})
                rtype = result.get("type", "")

                if rtype == "succeeded":
                    items.append(
                        BatchResultItem(
                            custom_id=custom_id,
                            success=True,
                            text=result.get("text") or None,
                            raw=rec,
                        )
                    )
                else:
                    items.append(
                        BatchResultItem(
                            custom_id=custom_id,
                            success=False,
                            error_message=result.get(
                                "error_message", "Unknown Azure error"
                            ),
                            raw=rec,
                        )
                    )

        logger.info(
            "Parsed %d Azure results (%d succeeded, %d failed)",
            len(items),
            sum(1 for i in items if i.success),
            sum(1 for i in items if not i.success),
        )
        return items


# ══════════════════════════════════════════════════════════════════════
# Provider registry
# ══════════════════════════════════════════════════════════════════════

# Maps provider name -> client class.  Add new providers here.
_PROVIDERS: Dict[str, type] = {
    "anthropic": AnthropicClient,
    "azure": AzureOpenAIClient,
}


def get_llm_client(provider: str = "") -> BaseLLMClient:
    """Factory function to get the appropriate LLM client.

    Args:
        provider: The provider name (must match a key in the registry).
            Currently supported: ``"anthropic"``, ``"azure"``.  If empty,
            falls back to ``config.LLM_PROVIDER``.

    Returns:
        An instance of the appropriate ``BaseLLMClient`` subclass.

    Raises:
        ValueError: If the provider is not recognised.

    Example:
        To add a new provider (e.g. Google Gemini)::

            class GeminiClient(BaseLLMClient):
                ...

            _PROVIDERS["gemini"] = GeminiClient

        Then set ``LLM_PROVIDER = "gemini"`` in ``config.py``.
    """
    if not provider:
        provider = config.LLM_PROVIDER
    provider = provider.lower().strip()
    client_class = _PROVIDERS.get(provider)

    if client_class is None:
        raise ValueError(
            f"Unknown LLM provider: '{provider}'. "
            f"Available: {list(_PROVIDERS.keys())}"
        )

    logger.info("Initialising LLM client: %s", provider)
    return client_class()
