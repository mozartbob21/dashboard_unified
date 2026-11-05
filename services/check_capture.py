"""Record the outcome of the original operation; never save its raw payload."""
from contextlib import contextmanager
from dataclasses import dataclass, field

from services import run_history


@dataclass
class CheckOutcome:
    status: str = 'success'
    metrics: dict = field(default_factory=dict)


@contextmanager
def capture(module, *, params=None, private_owner=None):
    run_id = run_history.record_start(module, params=params, private_owner=private_owner)
    outcome = CheckOutcome()
    try:
        yield outcome
    except Exception:
        run_history.record_finish(run_id, 'skipped' if outcome.status == 'skipped' else 'error')
        raise
    except BaseException:
        run_history.record_finish(run_id, 'interrupted')
        raise
    else:
        run_history.record_finish(run_id, outcome.status, metrics=outcome.metrics)
