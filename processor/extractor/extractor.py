from datetime import timedelta
from pathlib import Path
import json
import os
import pandas as pd

from config import logger, PROJECT_PATH


class CollectOperationExtractor:

    def __init__(self):
        self.collect_data = None
        self.logger = logger

    def retrieve_collects(self):
        """Retrieve and consolidate files"""
        collects = []
        filenames = os.listdir(Path(PROJECT_PATH / 'data'))
        for filename in filenames:
            if filename.endswith('.json'):
                file = Path(PROJECT_PATH / 'data' / filename)
                collects += json.load(file.open('r'))

        self.collect_data = collects
        return self.collect_data


