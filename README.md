# HARDWARIO Cloud Fetch

This repository contains a Python 3 CLI script to fetch device messages in HARDWARIO Cloud and extract useful measurements from it. The measurements are written to the specified XLSX file. It is not distributed via PyPI as code modification by customers is expected, and simplicity is the goal.


## Requirements

* Python 3


## Installation

1. Clone this repository:

       git clone git@github.com:hardwario/cloud-fetch.git

1. Install the Python dependencies:

       python3 -m pip install click pandas pendulum requests

> It is recommended to use `venv` for package isolation.


## Usage

1. See the script command line options:

       python3 fetch.py --help

1. Run the script:

       python3 fetch.py -g GROUP_ID -t GROUP_TOKEN

> The measurements can be easily modified in the `extract_measurement` function.


## Authors

* [**Pavel Hübner**](https://github.com/hubpav) - Initial work


## License

This project is licensed under the [**MIT License**](https://opensource.org/licenses/MIT/) - see the [**LICENSE**](https://github.com/hardwario/cloud-fetch/blob/master/LICENSE) file for details.
