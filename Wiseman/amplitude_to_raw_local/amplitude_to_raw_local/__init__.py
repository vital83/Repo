from dagster import Definitions, load_assets_from_modules

from . import fetching_assets, load_amplitude_assets, jobs

all_assets = load_assets_from_modules([load_amplitude_assets, fetching_assets])


defs = Definitions(
    assets=all_assets,
    schedules=[jobs.basic_schedule],
)
