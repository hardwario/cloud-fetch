import click
import json
import pandas
import pendulum
import requests
import sys

# Ohmic value of the load resistor
LOAD_RESISTOR = 39


def extract_measurement(message):

    if 'battery' not in message['data']:
        return None

    # Fetch message date
    dt = pendulum.parse(message['created_at'])
    dt = dt.in_timezone('Europe/Prague')
    date = dt.to_datetime_string()

    # Fetch voltages
    v1 = message['data']['battery'].get('voltage1')
    v2 = message['data']['battery'].get('voltage2')

    # Fetch temperatures
    t1 = message['data'].get('sensor', {}).get('thermometer', {}).get('temperature')
    t2 = message['data'].get('sensor', {}).get('hygrometer', {}).get('temperature')

    # If both voltages exist, create a measurement record
    if v1 and v2:
        return {
            'label': message['label'],
            'date': date,
            't1': round(t1, 1) if t1 is not None else None,
            't2': round(t2, 1) if t2 is not None else None,
            'v1': round(v1, 2),
            'v2': round(v2, 2),
            'r': round(LOAD_RESISTOR * (v1 - v2) / v2, 2)
        }

    return None


class FetchException(Exception):
    pass


# This class handles fetching of device list from the provided group
class DeviceFetcher:

    def __init__(self, group_id, api_token):
        self._group_id = group_id
        self._api_token = api_token

    def _get(self, limit=100, offset=0):
        params = {
            'group_id': self._group_id,
            'limit': limit,
            'offset': offset
        }
        headers = {
            'Authorization': 'Bearer {}'.format(self._api_token)
        }
        r = requests.get('https://api.hardwario.cloud/v1/devices', params=params, headers=headers)
        if r.status_code != 200:
            raise FetchException
        return json.loads(r.text)

    def fetch(self):
        found = False; records = []; offset = 0; limit = 100
        while True:
            found = True; click.echo('.', nl=False)
            data = self._get(limit=limit, offset=offset)
            count = len(data)
            records += data; offset += count
            if count == 0 or count < limit:
                break
        if found:
            click.echo()
        return records


# This class handles fetching of message list from the provided device
class MessageFetcher:

    def __init__(self, group_id, device_id, api_token):
        self._group_id = group_id
        self._device_id = device_id
        self._api_token = api_token

    def _get(self, limit=100, offset=0):
        params = {
            'group_id': self._group_id,
            'device_id': self._device_id,
            'limit': limit,
            'offset': offset
        }
        headers = {
            'Authorization': 'Bearer {}'.format(self._api_token)
        }
        r = requests.get('https://api.hardwario.cloud/v1/messages', params=params, headers=headers)
        if r.status_code != 200:
            raise FetchException
        return json.loads(r.text)

    def fetch(self):
        found = False; records = []; offset = 0; limit = 100
        while True:
            found = True; click.echo('.', nl=False)
            data = self._get(limit=limit, offset=offset)
            count = len(data)
            records += data; offset += count
            if count == 0 or count < limit:
                break
        if found:
            click.echo()
        return records


@click.command()
@click.option('--xlsx-file', '-x', metavar='XLSX_FILE', required=True, help='Specify output XLSX file.')
@click.option('--group-id', '-g', metavar='GROUP_ID', required=True, help='Provide group identifier.')
@click.option('--api-token', '-t', metavar='API_TOKEN', required=True, help='Provide group api token.')
def main(xlsx_file, group_id, api_token):

    # Fetch all devices in group
    click.echo('Fetching data for group: {}...'.format(group_id))
    devices = DeviceFetcher(group_id=group_id, api_token=api_token).fetch()

    # Initialize sheets
    sheets = []

    for device in devices:

        # Fetch all messages for device
        click.echo('Fetching messages for device: {} ({})...'.format(device['id'], device['name']))
        messages = MessageFetcher(device['group_id'], device['id'], device['api_token']).fetch()

        # Initialize measurements
        measurements = []

        # Walk through messages
        for m in messages:

            # Extract desired measurement out of the message
            measurement = extract_measurement(m)

            # Append the measurement if it was extracted
            if measurement is not None:
                measurements.append(measurement)

        # Append sheet as Pandas data frame
        df = pandas.DataFrame(measurements)
        sheets.append({'name': device['id'][:30], 'data': df})

    # Write measurements to XLSX
    click.echo('Generating XLSX file...')
    with pandas.ExcelWriter(xlsx_file) as writer:
        for sheet in sheets:
            sheet['data'].to_excel(writer, sheet_name=sheet['name'])

if __name__ == '__main__':
    try:
        main()
    except FetchException:
        click.echo('Request to REST API failed. Please, check your parameters!', err=True)
    except KeyboardInterrupt:
        pass
