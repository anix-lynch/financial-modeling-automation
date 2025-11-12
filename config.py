"""
Configuration for Financial Modeling Automation
"""

# Path to your existing models
VC_MODELS_PATH = "/Users/anixlynch/dev/northstar/vc fund model"

# Sample model parameters
DEFAULT_CAP_TABLE = {
    "fund_size": 8000000,
    "seed_investment": 2000000,
    "series_a_investment": 10000000,
    "series_b_investment": 30000000,
}

DEFAULT_FUND_MODEL = {
    "fund_size": 150000000,
    "management_fee_early": 0.02,
    "management_fee_late": 0.015,
    "carry_percentage": 0.20,
    "investment_period": 5,
    "fund_life": 10,
}

