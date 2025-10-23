from setuptools import setup, find_packages

setup(
    name="vbaProjectCompiler",
    packages=find_packages(),
    install_requires=[
        # Core dependencies will be installed from requirements.txt
    ],
    extras_require={
        "tests": [
            "pytest>=7.0",
            "pytest-cov>=4.0",
            "flake8>=6.0",
            "coveralls>=3.0",
        ],
    },
    tests_require=["pytest"],
)
