#!/bin/sh

set -eux

python3 fetch.py -x ${XLSX_FILE} -g ${GROUP_ID} -t ${API_TOKEN}
