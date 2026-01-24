from __future__ import annotations

from typing import Any

from .exact_match import exact_match


def _list_score(truth_list: list[Any], pred_list: Any) -> float:
    """Compute a partial match score for lists.

    Args:
        truth_list: Ground-truth list.
        pred_list: Predicted list.

    Returns:
        Fraction of truth elements present in prediction (order-insensitive).
    """
    if not isinstance(pred_list, list):
        return 0.0
    if not truth_list:
        return 0.0
    # Use exact match on elements; ignore order and duplicates.
    truth_set = {_normalize_scalar(v) for v in truth_list}
    pred_set = {_normalize_scalar(v) for v in pred_list}
    if not truth_set:
        return 0.0
    return len(truth_set & pred_set) / len(truth_set)


def _normalize_scalar(value: Any) -> str:
    """Normalize scalar values for set comparison."""
    if value is None:
        return "null"
    return str(value).strip()


def _list_score_ordered(truth_list: list[Any], pred_list: Any) -> float:
    """Compute an order-aware partial match score for lists.

    Args:
        truth_list: Ground-truth list.
        pred_list: Predicted list.

    Returns:
        LCS-based fraction of truth elements matched in order.
    """
    if not isinstance(pred_list, list):
        return 0.0
    if not truth_list:
        return 0.0
    truth_norm = [_normalize_scalar(v) for v in truth_list]
    pred_norm = [_normalize_scalar(v) for v in pred_list]
    lcs_len = _lcs_length(truth_norm, pred_norm)
    return lcs_len / len(truth_norm)


def _lcs_length(a: list[str], b: list[str]) -> int:
    """Compute the length of the longest common subsequence."""
    if not a or not b:
        return 0
    dp = [0] * (len(b) + 1)
    for i in range(1, len(a) + 1):
        prev = 0
        for j in range(1, len(b) + 1):
            temp = dp[j]
            if a[i - 1] == b[j - 1]:
                dp[j] = prev + 1
            else:
                dp[j] = max(dp[j], dp[j - 1])
            prev = temp
    return dp[-1]


def key_score(truth: Any, pred: Any) -> float:
    """Compute a key-level score against the truth payload.

    Args:
        truth: Ground-truth JSON payload.
        pred: Predicted JSON payload.

    Returns:
        Score in [0, 1]. For dict payloads, this is the fraction of truth keys
        that exactly match in the prediction. For non-dict payloads, this is
        1.0 if exactly equal, else 0.0.
    """
    if isinstance(truth, dict):
        total = len(truth)
        if total == 0:
            return 0.0
        if not isinstance(pred, dict):
            return 0.0
        score_sum = 0.0
        for key, truth_val in truth.items():
            if key not in pred:
                continue
            pred_val = pred[key]
            if isinstance(truth_val, list):
                score_sum += _list_score(truth_val, pred_val)
                continue
            score_sum += 1.0 if exact_match(truth_val, pred_val) else 0.0
        return score_sum / total
    if isinstance(truth, list):
        return _list_score(truth, pred)
    return 1.0 if exact_match(truth, pred) else 0.0


def key_score_ordered(truth: Any, pred: Any) -> float:
    """Compute a key-level score that respects list order."""
    if isinstance(truth, dict):
        total = len(truth)
        if total == 0:
            return 0.0
        if not isinstance(pred, dict):
            return 0.0
        score_sum = 0.0
        for key, truth_val in truth.items():
            if key not in pred:
                continue
            pred_val = pred[key]
            if isinstance(truth_val, list):
                score_sum += _list_score_ordered(truth_val, pred_val)
                continue
            score_sum += 1.0 if exact_match(truth_val, pred_val) else 0.0
        return score_sum / total
    if isinstance(truth, list):
        return _list_score_ordered(truth, pred)
    return 1.0 if exact_match(truth, pred) else 0.0
