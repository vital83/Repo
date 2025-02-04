from setuptools import find_packages, setup

setup(
    name="amplitude_to_raw_local",
    packages=find_packages(exclude=["amplitude_to_raw_local_tests"]),
    install_requires=[
        "dagster",
        "dagster-cloud"
    ],
    extras_require={"dev": ["dagster-webserver", "pytest"]},
)
