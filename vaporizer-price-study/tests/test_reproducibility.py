"""Regression tests for portability and published numerical results.

These tests RECOMPUTE quantities from the raw Terapeak CSVs; they do not read
cached expected outputs. Run with:  python -m pytest tests/ -v
"""
import os
import subprocess
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import pytest
from scipy import stats as sp_stats

PROJECT_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_DIR / 'data'

BAN_DATE = pd.Timestamp('2026-01-01')
PROPOSAL_DATE = pd.Timestamp('2022-04-05')
AGREEMENT_DATE = pd.Timestamp('2023-10-05')
ADOPTION_DATE = pd.Timestamp('2024-02-07')
EXCLUDE_KEYWORDS = ['key fill', 'keyed filler', 'bottle adapter', 'easy-fil',
                    'adapter', 'filler cap', 'pour fill', 'anti-spill',
                    'fill adapter', 'lot of 10']
VET_KEYWORDS = ['anesthesia system', 'anesthesia center', 'with cart',
                'rolling stand', 'w/ cart', 'rodent anesthesia',
                'compac5', 'rc2 rodent', 'v-10 isoflurane laboratory',
                'anesthesia workstation', 'apollo anesthesia']


def classify_period(d):
    d = pd.Timestamp(d)
    if d < PROPOSAL_DATE:
        return 1
    elif d < AGREEMENT_DATE:
        return 2
    elif d < ADOPTION_DATE:
        return 3
    elif d < BAN_DATE:
        return 4
    else:
        return 5


def load_terapeak(path):
    df = pd.read_csv(path)
    df['date_sold'] = pd.to_datetime(df['sold_date'], format='mixed',
                                     dayfirst=False, errors='coerce')
    df['price_usd'] = pd.to_numeric(df['price_usd'], errors='coerce')
    df = df.dropna(subset=['date_sold', 'price_usd'])
    return df


@pytest.fixture(scope='module')
def all_data():
    """Re-derive the analysis dataset from raw Terapeak inputs (mirrors analyze.py)."""
    frames = {}
    for agent, fname in [('Desflurane', 'terapeak_desflurane.csv'),
                         ('Sevoflurane', 'terapeak_sevoflurane.csv'),
                         ('Isoflurane', 'terapeak_isoflurane.csv')]:
        df = load_terapeak(DATA_DIR / fname)
        if 'item_category' in df.columns and agent != 'Desflurane':
            df = df[df['item_category'] == 'vaporizer']
        df = df.copy()
        mask = ~df['title'].str.lower().str.contains('|'.join(EXCLUDE_KEYWORDS), na=False)
        df = df[mask]
        if agent in ('Sevoflurane', 'Isoflurane'):
            mask = ~df['title'].str.lower().str.contains('|'.join(VET_KEYWORDS), na=False)
            df = df[mask]
        mask = ~df['title'].str.lower().str.contains(r'^lot of \d', na=False, regex=True)
        df = df[mask]
        df['agent_type'] = agent
        frames[agent] = df
    data = pd.concat(frames.values(), ignore_index=True)
    data['period'] = np.where(data['date_sold'] < BAN_DATE, 'pre', 'post')
    data['period_num'] = data['date_sold'].apply(classify_period)
    return data


def test_no_machine_specific_paths_in_scripts():
    for py in PROJECT_DIR.glob('*.py'):
        text = py.read_text()
        assert '/home/ubuntu/' not in text, f'{py.name} contains an absolute local path'
        assert 'vaporizer_research' not in text, f'{py.name} references the old external directory'


def test_data_files_discoverable():
    for f in ['terapeak_desflurane.csv', 'terapeak_sevoflurane.csv',
              'terapeak_isoflurane.csv', 'ebay_asking_prices.csv']:
        assert (DATA_DIR / f).is_file(), f'missing {f}'


def test_sample_counts(all_data):
    counts = all_data['agent_type'].value_counts()
    assert counts['Desflurane'] == 158
    assert counts['Sevoflurane'] == 321
    assert counts['Isoflurane'] == 554
    assert len(all_data) == 1033


def test_desflurane_pre_post(all_data):
    ad = all_data[all_data['agent_type'] == 'Desflurane']
    pre = ad[ad['period'] == 'pre']['price_usd']
    post = ad[ad['period'] == 'post']['price_usd']
    assert len(pre) == 141 and len(post) == 17
    assert pre.mean() == pytest.approx(238.30, abs=0.5)
    assert pre.std() == pytest.approx(194.75, abs=0.5)
    assert post.mean() == pytest.approx(163.58, abs=0.5)
    pct = (post.mean() - pre.mean()) / pre.mean() * 100
    assert pct == pytest.approx(-31.4, abs=0.1)
    _, tp = sp_stats.ttest_ind(pre, post, equal_var=False)
    _, up = sp_stats.mannwhitneyu(pre, post, alternative='two-sided')
    assert tp == pytest.approx(0.027, abs=0.001)
    assert up == pytest.approx(0.148, abs=0.001)
    pooled = np.sqrt((pre.var() + post.var()) / 2)
    assert (post.mean() - pre.mean()) / pooled == pytest.approx(-0.47, abs=0.01)


@pytest.mark.parametrize('agent,rho,rho_p,tau,tau_p', [
    ('Desflurane', -0.277, 0.0005, -0.121, 0.049),
    ('Sevoflurane', 0.010, 0.862, 0.039, 0.365),
    ('Isoflurane', -0.086, 0.044, -0.074, 0.025),
])
def test_trend_tests_use_correct_predictors(all_data, agent, rho, rho_p, tau, tau_p):
    ad = all_data[all_data['agent_type'] == agent].copy()
    # Spearman: individual sale DATE (ordinal) vs price
    r, rp = sp_stats.spearmanr(ad['date_sold'].map(lambda d: d.toordinal()),
                             ad['price_usd'])
    # Kendall tau-b: ordinal regulatory phase (1-5) vs price, NOT the date
    t, tp = sp_stats.kendalltau(ad['period_num'], ad['price_usd'])
    assert r == pytest.approx(rho, abs=0.005)
    assert rp == pytest.approx(rho_p, abs=max(rho_p * 0.05, 0.001))
    assert t == pytest.approx(tau, abs=0.005)
    assert tp == pytest.approx(tau_p, abs=max(tau_p * 0.05, 0.001))


@pytest.mark.parametrize('agent,q_rho,q_p', [
    ('Desflurane', -0.601, 0.039),
    ('Sevoflurane', -0.092, 0.777),
    ('Isoflurane', -0.321, 0.285),
])
def test_quarterly_median_trend(all_data, agent, q_rho, q_p):
    ad = all_data[all_data['agent_type'] == agent]
    qm = ad.groupby(ad['date_sold'].dt.to_period('Q'))['price_usd'].agg(['median', 'count'])
    qm = qm[qm['count'] >= 3]
    r, p = sp_stats.spearmanr(range(len(qm)), qm['median'].values)
    assert r == pytest.approx(q_rho, abs=0.005)
    assert p == pytest.approx(q_p, abs=max(q_p * 0.05, 0.002))


def test_scripts_run_from_external_cwd(tmp_path):
    """analyze.py and create_jca_tables.py must run from outside the repo."""
    for script in ['analyze.py', 'create_jca_tables.py']:
        res = subprocess.run([sys.executable, str(PROJECT_DIR / script)],
                             cwd=str(tmp_path), capture_output=True, text=True,
                             timeout=600)
        assert res.returncode == 0, f'{script} failed from external cwd:\n{res.stderr[-2000:]}'
