"""Retry helper for transient Azure/network failures.

Built on `tenacity` (already a project dependency) instead of hand-rolled
sleep loops, but wrapped in one small function so call sites stay readable
and the retry policy (attempts, backoff, which errors are retryable) lives
in one place.
"""

from __future__ import annotations

import logging
from typing import Callable, TypeVar

from azure.core.exceptions import ServiceRequestError, ServiceResponseError
from openai import APIConnectionError, APITimeoutError, InternalServerError, RateLimitError
from tenacity import (
    before_sleep_log,
    retry,
    retry_if_exception_type,
    stop_after_attempt,
    wait_exponential,
)
from tenacity.wait import wait_base


T = TypeVar("T")

# Errors worth retrying: rate limits, timeouts, transient connectivity/server
# issues. NOT retried: auth errors, bad requests, validation errors - retrying
# those just wastes time and hides a real bug/misconfiguration.
RECOVERABLE_LLM_ERRORS = (RateLimitError, APITimeoutError, APIConnectionError, InternalServerError)
RECOVERABLE_NETWORK_ERRORS = (TimeoutError, ConnectionError, ServiceRequestError, ServiceResponseError)


DEFAULT_MAX_RETRIES = 3
DEFAULT_BACKOFF_BASE_SECONDS = 2.0
LLM_MAX_RETRIES = 3
LLM_MAX_WAIT_SECONDS = 120


def _retry_after_seconds(exc: BaseException | None) -> float | None:
    """Read the server-provided `Retry-After` header off a RateLimitError, if present."""
    response = getattr(exc, "response", None)
    headers = getattr(response, "headers", None)
    value = headers.get("retry-after") if headers else None
    if value is None:
        return None
    try:
        return max(float(value), 0.0)
    except (TypeError, ValueError):
        return None


class _RateLimitAwareWait(wait_base):
    """Honor the API's `Retry-After` header on 429s; fall back to exponential backoff
    for everything else (timeouts, connection errors, etc. don't carry that header).
    """

    def __init__(self, multiplier: float, max_seconds: float) -> None:
        self._exponential = wait_exponential(multiplier=multiplier, min=0, max=max_seconds)

    def __call__(self, retry_state) -> float:
        exc = retry_state.outcome.exception() if retry_state.outcome else None
        retry_after = _retry_after_seconds(exc)
        if retry_after is not None:
            return retry_after
        return self._exponential(retry_state)


def retry_on(*exception_types: type[BaseException], max_attempts: int | None = None):
    """Decorator factory: retry a function on the given exception types only.

    Retry count and backoff come from Settings (SSI_MAX_RETRIES,
    SSI_BACKOFF_BASE_SECONDS) so tuning them doesn't require a code change.
    """
    attempts = max_attempts if max_attempts is not None else DEFAULT_MAX_RETRIES

    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        return retry(
            reraise=True,
            stop=stop_after_attempt(attempts),
            wait=wait_exponential(multiplier=DEFAULT_BACKOFF_BASE_SECONDS, min=0, max=60),
            retry=retry_if_exception_type(exception_types),
            # before_sleep=before_sleep_log(logger, logging.WARNING),
        )(func)

    return decorator


def _retry_llm(func: Callable[..., T]) -> Callable[..., T]:
    """LLM-specific retry: honors 429 `Retry-After` headers and allows a larger
    retry budget than generic network calls, since rate limits often need a
    real wait rather than a quick exponential backoff.
    """
    return retry(
        reraise=True,
        stop=stop_after_attempt(LLM_MAX_RETRIES),
        wait=_RateLimitAwareWait(multiplier=DEFAULT_BACKOFF_BASE_SECONDS, max_seconds=LLM_MAX_WAIT_SECONDS),
        retry=retry_if_exception_type(RECOVERABLE_LLM_ERRORS),
        # before_sleep=before_sleep_log(logger, logging.WARNING),
    )(func)


# Ready-made decorators for the two external calls this pipeline makes.
retry_doc_intel_call = retry_on(*RECOVERABLE_NETWORK_ERRORS)
retry_llm_call = _retry_llm
