import json

import pandas as pd
from pandas import DataFrame
from datetime import datetime

from config import PROJECT_PATH
from config.definitions import CollectFields


class CollectTransformer:
    """This class aims at structuring a list of dictionaries describing collect operations."""
    COLS_TO_KEEP = [
        CollectFields.updated_at,
        CollectFields.website,
        CollectFields.territory_uid,
        CollectFields.id,
        CollectFields.item_scraped_count,
        CollectFields.finish_reason,
        CollectFields.status,
    ]

    TERRITORIES_INFOS_PATH = PROJECT_PATH / 'metadata' / 'commune_epci_departement.tsv'

    DEPARTMENTS_INFOS_PATH = PROJECT_PATH / 'metadata' / 'departements.json'

    def __init__(self):
        self.collect_data = None
        self.departements_infos = self.get_departements_infos()
        self.territories_infos = self.get_territories_infos()

    def load_and_transform(self, data):
        self.collect_data = self.load(data)
        self.collect_data[CollectFields.updated_at] = pd.to_datetime(self.collect_data[CollectFields.updated_at])
        return self.transform(self.collect_data)

    @staticmethod
    def load(data):
        return pd.DataFrame(data)

    @classmethod
    def transform(cls, collects: DataFrame) -> DataFrame:
        collects[CollectFields.item_scraped_count] = collects.apply(
            lambda x: cls.get_items_scraped_count(x), axis=1
        )
        collects[CollectFields.finish_reason] = collects.apply(
            lambda x: cls.get_finish_reason(x), axis=1
        )

        collects = collects.filter(items=cls.COLS_TO_KEEP)

        return collects

    @staticmethod
    def get_items_scraped_count(x):
        try:
            return x["infos"]["stats"]["item_scraped_count"]
        except KeyError:
            return 0  # No item scraped count attribute, return 0

    @staticmethod
    def get_finish_reason(x):
        try:
            return x["infos"]["stats"]["finish_reason"]
        except KeyError:
            return 0

    @staticmethod
    def add_territories_info(collects):
        """Keeps the most recent collect when there is more than one by territory that passed the check and renames
        relevant columns"""

        collects['code commune'] = collects['territory_uid'].apply(lambda x: x[6:])
        formatted_collects = collects.merge(self.territories_infos, on='code commune')

        formatted_collects["departement"] = formatted_collects["territory_uid"].apply(
            lambda x: self.departements_infos.get(x[6:8]) if x[:6] != 'FREPCI' else None)

        formatted_collects["Id de l'alerte"] = ""  # Colonne remplie par Eric
        formatted_collects['Date de mise à jour'] = datetime.today().strftime('%Y-%m-%d')

        formatted_collects.rename(columns={
            'departement': 'Nom département',
            'territory_uid': 'Code',
            'nom commune': 'Nom',
            'website': 'URL dans Pensieve',
            'finish_reason': 'Statut de la collecte'
        }, inplace=True)

        return formatted_collects

    @classmethod
    def get_departements_infos(cls):
        departements_infos = json.load(cls.DEPARTMENTS_INFOS_PATH.open('r'))
        return departements_infos

    @classmethod
    def get_territories_infos(cls):
        territories_infos = pd.read_csv(cls.TERRITORIES_INFOS_PATH, sep='\t')
        return territories_infos[['code commune', 'nom commune']]
