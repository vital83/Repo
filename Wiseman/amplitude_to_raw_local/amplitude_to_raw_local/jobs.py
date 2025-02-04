from dagster import AssetSelection, ScheduleDefinition, define_asset_job

asset_job = define_asset_job("asset_job",  selection="amplitude_zip_file*")

basic_schedule = ScheduleDefinition(job=asset_job, cron_schedule="20 21 * * *", execution_timezone="Europe/Moscow")