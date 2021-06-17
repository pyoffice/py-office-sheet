import pathlib
from setuptools import setup

# The directory containing this file
HERE = pathlib.Path(__file__).parent

# The text of the README file
README = (HERE / "README.md").read_text()

# This call to setup() does all the work
setup(
    name="py-office-sheet",
    version="0.1.195",
    description="py-office-sheet is a gui based on pandas and numpy to process data efficienly",
    long_description=README,
    long_description_content_type="text/markdown",
    url="https://github.com/YC-Lammy/py-office-sheet",
    author="YC",
    author_email="yclam508@protonmail.com",
    platforms = "Linux, Mac OS X, Windows",
    license="GPL",
    classifiers=[
        "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
        "Programming Language :: Python :: 3",
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3 :: Only',
    ],
    packages=["pyOfficeSheet"],
    include_package_data=True,
    install_requires=["joblib", "PySide2","pandas","numpy","matplotlib","pyshortcuts"],
    entry_points={
        'console_scripts': [
            'py-office-sheet=pyOfficeSheet.__main__:main',],
    },
)
