"""Numerical checks of the forward year-end-2026 valuation, in USD millions.

2027-2035 cash flows are discounted to Dec 31, 2026. The bridge rolls reported
Sep-2026 cash forward with a simplified Q4 FCF estimate. This is not a value
discounted to the November benchmark quote date.
"""
import math


def forecast(a, p, uplift=0., sbc_addback=0.):
    if any(len(a[k]) != 10 for k in ['proc_g','iapp_g','place_g','svc_g','sga']):
        raise ValueError('Exactly ten annual assumptions are required')
    proc=p['procedures'];iapp=p['instruments']*1e6/proc
    placements=p['placements'];rpp=p['systems']*1e6/placements
    installed=p['installed'];spb=p['services']*1e6/((p['installed_prior']+installed)/2)
    previous=p['revenue'];rows=[]
    for i in range(10):
        if a['proc_g'][i]+uplift<=-1:raise ValueError('Procedure growth must exceed -100%')
        proc*=1+a['proc_g'][i]+uplift;iapp*=1+a['iapp_g'][i]
        placements*=1+a['place_g'][i];spb*=1+a['svc_g'][i]
        closing=installed+placements*a['net_add']
        services=spb*(installed+closing)/2/1e6;installed=closing
        revenue=proc*iapp/1e6+placements*rpp/1e6+services
        ebit=revenue*(a['gm']-a['rd']-a['sga'][i])
        fcf=ebit*(1-a['tax'])+revenue*a['da']-revenue*a['capex']-(revenue-previous)*a['nwc']+sbc_addback*revenue
        rows.append({'year':2026+i,'revenue':revenue,'procedures':proc,'ebit':ebit,'fcf':fcf})
        previous=revenue
    return rows


def value(a,p,uplift=0.,sbc_addback=0.,wacc=None,tg=None):
    wacc=a['wacc'] if wacc is None else wacc;tg=a['tg'] if tg is None else tg
    if not all(math.isfinite(x) for x in [wacc,tg]) or wacc<=tg or wacc<=-1:
        raise ValueError('Finite WACC must exceed terminal growth and -100%')
    if p['shares']<=0:raise ValueError('Shares outstanding must be positive')
    rows=forecast(a,p,uplift,sbc_addback)
    pv=sum(row['fcf']/(1+wacc)**i for i,row in enumerate(rows[1:],start=1))
    terminal=rows[-1]['fcf']*(1+tg)/(wacc-tg)/(1+wacc)**9
    bridge=.25*rows[0]['fcf']+p['cash']-p['settlement']-p['nci']
    return (pv+terminal+bridge)/p['shares'],rows


def solve_growth_uplift(a,p,price,low=-.2,high=.5):
    if not math.isfinite(price) or price<=0:raise ValueError('Benchmark must be finite and positive')
    if not value(a,p,low)[0]<=price<=value(a,p,high)[0]:
        raise ValueError('Benchmark is outside the solver bracket')
    for _ in range(80):
        mid=(low+high)/2
        if value(a,p,mid)[0]<price:low=mid
        else:high=mid
    return (low+high)/2
