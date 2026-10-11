from setuptools import setup, find_packages

setup(
    name="jago-docx",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        "python-docx>=1.1.0",
        "lxml>=5.0.0"
    ],
    entry_points={
        "console_scripts": [
            "jago-docx=jago_docx.cli:main",
        ],
    },
)
