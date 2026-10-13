"""Chat-specific limits; provider and TLS configuration stay in the shared client."""
import math
import os

import httpx


def model_timeout():
    """Allow a long final answer without extending connection or upload waits.

    AICHAT_MODEL_READ_TIMEOUT_SECONDS is a read-inactivity limit, not a total
    request deadline. It defaults to 15 minutes and is bounded to 1–30 minutes.
    Invalid/non-positive values use the default. Other AI modules and the
    short skill-selection calls retain their existing timeouts.
    """
    try:
        seconds = float(os.environ.get('AICHAT_MODEL_READ_TIMEOUT_SECONDS', '900'))
    except (TypeError, ValueError):
        seconds = 900.0
    if not math.isfinite(seconds) or seconds <= 0:
        seconds = 900.0
    return httpx.Timeout(connect=15.0, read=min(1800.0, max(60.0, seconds)),
                         write=30.0, pool=15.0)
