from setuptools import setup, find_packages

setup(
    name="vbaProjectCompiler",
    packages=find_packages(),
    install_requires=[
        "ms_ovba_compression",
        "ms_ovba_crypto",
        "ms_cfb @ git+https://github.com/James-McNeil/MS-CFB@main",
        "ms_pcode_assembler @ git+https://github.com/James-McNeil/MS-Pcode-Assembler@main",
    ],
    extras_require={
        "tests": [
            "pytest>=7.0",
            "pytest-cov>=4.0",
            "flake8>=6.0",
            "flake8-annotations",
            "mypy",
            "pep8-naming",
            "coveralls>=3.0",
        ],
    },
    tests_require=["pytest"],
)
