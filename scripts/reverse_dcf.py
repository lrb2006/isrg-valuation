"""Reproduce forward valuation and reverse DCF. No workbook writes on import."""
import json
from pathlib import Path
from build_model import SCENARIOS, VALUATION_INPUTS, PRICE, KEY_DEBATE
from valuation import value, solve_growth_uplift


def results():
    a=SCENARIOS['Base'];p=VALUATION_INPUTS
    uplift=solve_growth_uplift(a,p,PRICE)
    rows=value(a,p,uplift)[1]
    return {'valuation_date':'2026-12-31 (forward value)',
            'benchmark_price':PRICE,'benchmark_verified':False,
            'scenarios':{name:value(a,p)[0] for name,a in SCENARIOS.items()},
            'reverse_uplift':uplift,'reverse_revenue_cagr':(rows[-1]['revenue']/p['revenue'])**.1-1,
            'sensitivity':dict(KEY_DEBATE)}


if __name__=='__main__':
    output=results()
    folder=Path(__file__).resolve().parents[1]/'reports';folder.mkdir(exist_ok=True)
    (folder/'valuation_results.json').write_text(json.dumps(output,indent=2))
    print(json.dumps(output,indent=2))
