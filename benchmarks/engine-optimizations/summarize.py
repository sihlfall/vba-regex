from collections import defaultdict
import json
import math
from pathlib import Path
from statistics import median
import sys


def geometric_mean(values):
    return math.exp(sum(math.log(value) for value in values) / len(values))


def summarize(folder):
    samples = json.loads((folder / 'samples.json').read_text(encoding='utf-8-sig'))
    groups = defaultdict(list)
    for row in samples:
        groups[(row['case_id'], row['operation'], row['variant'])].append(row['ms'])
    comparisons = []
    for case in dict.fromkeys(row['case_id'] for row in samples):
        label = next(row['case'] for row in samples if row['case_id'] == case)
        for op in dict.fromkeys(row['operation'] for row in samples):
            before, after = groups[(case, op, 'baseline')], groups[(case, op, 'candidate')]
            comparisons.append({'case_id': case, 'case': label, 'operation': op,
                                'baseline_ms': median(before), 'candidate_ms': median(after),
                                'relative_time_percent': 100 * median(after) / median(before),
                                'baseline_range_ms': [min(before), max(before)],
                                'candidate_range_ms': [min(after), max(after)]})
    original = [item for item in comparisons if item['case_id'] < 6 and item['operation'] in ('List','ListArr','MatchAll','MatchAllArr')]
    result = {
        'original_suite_relative_time_percent': geometric_mean([item['relative_time_percent'] for item in original]),
        'expanded_suite_relative_time_percent': geometric_mean([item['relative_time_percent'] for item in comparisons]),
        'operation_relative_time_percent': {op: geometric_mean([item['relative_time_percent'] for item in comparisons if item['operation'] == op]) for op in dict.fromkeys(row['operation'] for row in samples)},
        'comparisons': comparisons,
    }
    (folder / 'summary.json').write_text(json.dumps(result, indent=2) + '\n')
    print('Original suite normalized time: %.2f%%; expanded suite: %.2f%% (baseline = 100%%).' % (result['original_suite_relative_time_percent'], result['expanded_suite_relative_time_percent']))
    return result


if __name__ == '__main__':
    summarize(Path(sys.argv[1]))
