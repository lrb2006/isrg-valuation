import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
import pytest
from build_model import SCENARIOS, VALUATION_INPUTS
from valuation import value, solve_growth_uplift


def test_source_revenue_components_reconcile():
    p=VALUATION_INPUTS
    assert p['instruments']==pytest.approx(6018.9)
    assert p['instruments']+p['systems']+p['services']==pytest.approx(p['revenue'])


@pytest.mark.parametrize('uplift',[-.02,0,.05])
def test_reverse_solver_recovers_known_growth(uplift):
    price=value(SCENARIOS['Base'],VALUATION_INPUTS,uplift)[0]
    assert solve_growth_uplift(SCENARIOS['Base'],VALUATION_INPUTS,price)==pytest.approx(uplift,abs=1e-10)


def test_larger_discount_rate_reduces_value():
    a=SCENARIOS['Base'];p=VALUATION_INPUTS
    assert value(a,p,wacc=.10)[0]<value(a,p,wacc=.08)[0]


def test_unbounded_terminal_value_is_rejected():
    with pytest.raises(ValueError,match='WACC'):value(SCENARIOS['Base'],VALUATION_INPUTS,wacc=.03,tg=.04)


def test_unreachable_reverse_target_does_not_return_false_solution():
    with pytest.raises(ValueError,match='bracket'):solve_growth_uplift(SCENARIOS['Base'],VALUATION_INPUTS,1e9)


def test_scenarios_are_ordered():
    p=VALUATION_INPUTS
    assert value(SCENARIOS['Bear'],p)[0]<value(SCENARIOS['Base'],p)[0]<value(SCENARIOS['Bull'],p)[0]
    rows=value(SCENARIOS['Base'],p)[1]
    assert len(rows)==10 and rows[0]['year']==2026 and rows[-1]['year']==2035
    assert all(row['fcf']>0 for row in rows)
