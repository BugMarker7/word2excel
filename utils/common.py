import os
from datetime import datetime


def get_output_file(output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    time_str = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    file_name = f'output_{time_str}.xlsx'
    return os.path.join(output_dir, file_name)
