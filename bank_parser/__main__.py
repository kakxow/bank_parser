import sys

from .aggregate_data import aggregate_data
from .process_data import process_result
from .utils import send_report


only_today = bool(int(sys.argv[1])) if len(sys.argv) > 1 else True
# Mail.
save_location = r'C:\Max\temp_for_registries'
result = aggregate_data(only_today, save_location)
report_path = process_result(result)

send_report(report_path, ['lgalkina@maxus.ru'])
