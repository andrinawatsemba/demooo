def largest_remainder_percentages(values):
    """Round a list of values to 1-decimal percentages that always sum
    to exactly 100.0. Fixes the old bug where each percentage was
    rounded independently (e.g. 6 categories each rounded to 1 decimal
    landing at 99.5% or 100.3% instead of 100.0%).

    Shared by app.py and pdf_report.py so there's one copy of this
    logic, not three slightly-different ones."""
    total = sum(values)
    if total <= 0:
        return [0.0 for _ in values]

    raw = [v / total * 100 for v in values]
    floored = [int(x * 10) / 10 for x in raw]
    remainders = [x * 10 - int(x * 10) for x in raw]

    diff_tenths = round((100 - sum(floored)) * 10)
    order = sorted(range(len(values)), key=lambda i: remainders[i], reverse=True)
    result = list(floored)
    for i in range(diff_tenths):
        result[order[i % len(order)]] = round(result[order[i % len(order)]] + 0.1, 1)
    return [round(r, 1) for r in result]
