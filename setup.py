from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="excelMCP",
    version="0.1.0",
    author="ExcelMCP Team",
    author_email="example@example.com",
    description="A library for working with Excel files and data manipulation",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/Sanjeev160489/mcp_repo",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
    install_requires=[
        "openpyxl>=3.1.0",
        "pandas>=1.5.3",
        "numpy>=1.24.0",
        "xlrd>=2.0.1",
        "xlsxwriter>=3.0.8",
    ],
)