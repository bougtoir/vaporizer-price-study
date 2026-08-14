"""
Sensitivity analyses for the JATM major revision.

Outputs:
  - data/revision_sensitivity.json
  - data/revision_sensitivity_its.csv

Includes:
  - Bootstrap confidence intervals for pre-/post-ban mean and median differences.
  - Post-hoc power estimates for Mann-Whitney U and Welch t-tests (lognormal simulation).
  - Difference-in-differences (DiD) on log-transformed transaction prices.
  - Comparative interrupted time-series (CITS) on monthly median prices.
"""
import json
import os
import warnings
import numpy as np
import pandas as pd
from scipy import stats
import statsmodels.api as sm
import statsmodels.formula.api as smf

warnings.filterwarnings('ignore')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, 'data')
BAN_DATE = pd.Timestamp('2026-01-01')
N_BOOT = 10000
N_SIM = 10000
RNG = np.random.default_rng(20260729)


def load_data():
    df = pd.read_csv(os.path.join(DATA_DIR, 'combined_cleaned.csv'))
    df['date_sold'] = pd.to_datetime(df['date_sold'])
    df = df[df['price_usd'] > 0].copy()
    df['post'] = (df['date_sold'] >= BAN_DATE).astype(int)
    df['months_since_start'] = (df['date_sold'].dt.year - df['date_sold'].min().year) * 12 + (
        df['date_sold'].dt.month - df['date_sold'].min().month)
    df['log_price'] = np.log(df['price_usd'])
    df['agent_desflurane'] = (df['agent_type'] == 'Desflurane').astype(int)
    df['agent_isoflurane'] = (df['agent_type'] == 'Isoflurane').astype(int)
    df['agent_sevoflurane'] = (df['agent_type'] == 'Sevoflurane').astype(int)
    return df


def bootstrap_pre_post(df, agent, n=N_BOOT):
    sub = df[df['agent_type'] == agent]
    pre = sub[sub['post'] == 0]['price_usd'].values
    post = sub[sub['post'] == 1]['price_usd'].values
    if len(pre) == 0 or len(post) == 0:
        return None
    diffs_mean = []
    diffs_median = []
    pvals_u = []
    pvals_t = []
    for _ in range(n):
        b_pre = RNG.choice(pre, size=len(pre), replace=True)
        b_post = RNG.choice(post, size=len(post), replace=True)
        diffs_mean.append(b_post.mean() - b_pre.mean())
        diffs_median.append(np.median(b_post) - np.median(b_pre))
        try:
            _, p_u = stats.mannwhitneyu(b_pre, b_post, alternative='two-sided')
        except ValueError:
            p_u = np.nan
        try:
            _, p_t = stats.ttest_ind(b_pre, b_post, equal_var=False)
        except ValueError:
            p_t = np.nan
        pvals_u.append(p_u)
        pvals_t.append(p_t)
    diffs_mean = np.array(diffs_mean)
    diffs_median = np.array(diffs_median)
    pvals_u = np.array(pvals_u)
    pvals_t = np.array(pvals_t)
    return {
        'n_pre': len(pre),
        'n_post': len(post),
        'observed_mean_diff': float(post.mean() - pre.mean()),
        'observed_median_diff': float(np.median(post) - np.median(pre)),
        'mean_diff_ci95': [float(np.percentile(diffs_mean, 2.5)),
                           float(np.percentile(diffs_mean, 97.5))],
        'median_diff_ci95': [float(np.percentile(diffs_median, 2.5)),
                             float(np.percentile(diffs_median, 97.5))],
        'prop_mean_diff_lt0': float(np.mean(diffs_mean < 0)),
        'prop_median_diff_lt0': float(np.mean(diffs_median < 0)),
        'mannwhitney_p_median': float(np.nanmedian(pvals_u)),
        'welch_p_median': float(np.nanmedian(pvals_t)),
    }


def fit_lognorm_params(x):
    # fit lognormal by MLE on log(x)
    logx = np.log(x)
    mu = logx.mean()
    sigma = logx.std(ddof=1)
    return mu, sigma


def power_simulation(pre, post, n_sim=N_SIM):
    """Estimate power using lognormal simulation with observed distributional parameters."""
    if len(pre) == 0 or len(post) == 0:
        return None
    mu_pre, sigma_pre = fit_lognorm_params(pre)
    mu_post, sigma_post = fit_lognorm_params(post)
    n_pre, n_post = len(pre), len(post)
    pow_u = 0.0
    pow_t = 0.0
    for _ in range(n_sim):
        sim_pre = RNG.lognormal(mu_pre, sigma_pre, size=n_pre)
        sim_post = RNG.lognormal(mu_post, sigma_post, size=n_post)
        _, p_u = stats.mannwhitneyu(sim_pre, sim_post, alternative='two-sided')
        _, p_t = stats.ttest_ind(sim_pre, sim_post, equal_var=False)
        if p_u < 0.05:
            pow_u += 1
        if p_t < 0.05:
            pow_t += 1
    return {
        'n_pre': n_pre,
        'n_post': n_post,
        'power_mannwhitney': float(pow_u / n_sim),
        'power_welch_t': float(pow_t / n_sim),
        'simulated_under': 'lognormal with observed group-specific mu, sigma',
    }


def did_transaction_level(df):
    """Difference-in-differences on log price, using sevoflurane as reference control."""
    df = df.copy()
    df['post_desflurane'] = df['post'] * df['agent_desflurane']
    df['post_isoflurane'] = df['post'] * df['agent_isoflurane']
    df['trend_desflurane'] = df['months_since_start'] * df['agent_desflurane']
    df['trend_isoflurane'] = df['months_since_start'] * df['agent_isoflurane']
    formula = (
        'log_price ~ C(agent_type) + months_since_start + post + '
        'post_desflurane + post_isoflurane + '
        'trend_desflurane + trend_isoflurane'
    )
    try:
        model = smf.ols(formula, data=df).fit(cov_type='HC3')
        return {
            'nobs': int(model.nobs),
            'r2': float(model.rsquared),
            'desflurane_post_coef': float(model.params.get('post_desflurane', np.nan)),
            'desflurane_post_p': float(model.pvalues.get('post_desflurane', np.nan)),
            'isoflurane_post_coef': float(model.params.get('post_isoflurane', np.nan)),
            'isoflurane_post_p': float(model.pvalues.get('post_isoflurane', np.nan)),
            'desflurane_trend_p': float(model.pvalues.get('trend_desflurane', np.nan)),
            'isoflurane_trend_p': float(model.pvalues.get('trend_isoflurane', np.nan)),
            'overall_post_coef': float(model.params.get('post', np.nan)),
            'overall_post_p': float(model.pvalues.get('post', np.nan)),
        }
    except Exception as e:
        return {'error': str(e)}


def monthly_panel(df):
    df['month'] = df['date_sold'].dt.to_period('M')
    monthly = df.groupby(['agent_type', 'month'])['price_usd'].agg(['median', 'count']).reset_index()
    monthly['month_start'] = monthly['month'].dt.to_timestamp()
    monthly['time_months'] = ((monthly['month_start'].dt.year - monthly['month_start'].min().year) * 12 +
                              (monthly['month_start'].dt.month - monthly['month_start'].min().month))
    monthly['post'] = (monthly['month_start'] >= BAN_DATE).astype(int)
    monthly['time_after'] = (monthly['month_start'] - BAN_DATE).dt.days / 30.44 * monthly['post']
    monthly['log_median'] = np.log(monthly['median'])
    return monthly


def cits_monthly_medians(monthly):
    """Comparative interrupted time-series on monthly median prices."""
    monthly = monthly.copy()
    monthly['agent_desflurane'] = (monthly['agent_type'] == 'Desflurane').astype(int)
    monthly['agent_isoflurane'] = (monthly['agent_type'] == 'Isoflurane').astype(int)
    monthly['post_desflurane'] = monthly['post'] * monthly['agent_desflurane']
    monthly['post_isoflurane'] = monthly['post'] * monthly['agent_isoflurane']
    monthly['trend_after_desflurane'] = monthly['time_after'] * monthly['agent_desflurane']
    monthly['trend_after_isoflurane'] = monthly['time_after'] * monthly['agent_isoflurane']
    formula = (
        'log_median ~ C(agent_type) + time_months + post + time_after + '
        'post_desflurane + post_isoflurane + '
        'trend_after_desflurane + trend_after_isoflurane'
    )
    try:
        # HAC robust standard errors to account for autocorrelation in time series
        model = smf.ols(formula, data=monthly).fit(cov_type='HAC', cov_kwds={'maxlags': 3})
        return {
            'nobs': int(model.nobs),
            'r2': float(model.rsquared),
            'desflurane_level_change_coef': float(model.params.get('post_desflurane', np.nan)),
            'desflurane_level_change_p': float(model.pvalues.get('post_desflurane', np.nan)),
            'desflurane_slope_change_coef': float(model.params.get('trend_after_desflurane', np.nan)),
            'desflurane_slope_change_p': float(model.pvalues.get('trend_after_desflurane', np.nan)),
            'isoflurane_level_change_coef': float(model.params.get('post_isoflurane', np.nan)),
            'isoflurane_level_change_p': float(model.pvalues.get('post_isoflurane', np.nan)),
        }
    except Exception as e:
        return {'error': str(e)}


def main():
    df = load_data()
    out = {}

    # Bootstrap and power
    bootstrap = {}
    power = {}
    for agent in ['Desflurane', 'Sevoflurane', 'Isoflurane']:
        sub = df[df['agent_type'] == agent]
        pre = sub[sub['post'] == 0]['price_usd'].values
        post = sub[sub['post'] == 1]['price_usd'].values
        bootstrap[agent] = bootstrap_pre_post(df, agent)
        power[agent] = power_simulation(pre, post)
    out['bootstrap_pre_post'] = bootstrap
    out['power_simulation'] = power

    # DiD / CITS
    out['did_transaction_level'] = did_transaction_level(df)
    monthly = monthly_panel(df)
    out['cits_monthly_medians'] = cits_monthly_medians(monthly)

    out_path = os.path.join(DATA_DIR, 'revision_sensitivity.json')
    with open(out_path, 'w') as f:
        json.dump(out, f, indent=2)
    print(f'Saved {out_path}')

    # Save monthly panel for diagnostics/figures
    monthly.to_csv(os.path.join(DATA_DIR, 'revision_sensitivity_its.csv'), index=False)
    print('Saved data/revision_sensitivity_its.csv')

    # Print human summary
    print(json.dumps(out, indent=2))


if __name__ == '__main__':
    main()
