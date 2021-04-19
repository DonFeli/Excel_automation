from datetime import timedelta

from config.definitions import CollectFields


class StoppedCollectDetector:
    def __init__(self, collects):
        self.collect_data = self.filter_insufficient_collects(collects, 2)
        self.pairs = self.get_last_collect_by_territory(collects)
        self.processed_pairs = self.get_processed_pairs(self.pairs)

    @staticmethod
    def filter_insufficient_collects(collect_data, minimal_collect_count):
        """Exclude collects from territories where the total number of collect is less than minimal_collect_count."""
        return collect_data.groupby(
            CollectFields.territory_uid
        ).filter(lambda x: x[CollectFields.id].count() >= minimal_collect_count)

    @staticmethod
    def get_last_collect_by_territory(collect_data):
        """Get the last available collect on a give"""
        last_collects = collect_data.loc[collect_data.groupby(['territory_uid'])["updated_at"].idxmax()]
        pairs = collect_data.merge(last_collects, on=[CollectFields.territory_uid], suffixes=('', '_last'))
        pairs.drop(['website_last', 'id_last', 'finish_reason_last', 'status_last'], axis=1, inplace=True)
        return pairs

    @staticmethod
    def get_processed_pairs(pairs):
        """Create a dataframe with, for each territory, combine the last collect with all other collects."""
        pairs["interval"] = pairs.apply(lambda x: x["updated_at_last"] - x["updated_at"], axis=1)
        pairs["is_stopped"] = pairs.apply(lambda x: x["item_scraped_count_last"] == 0, axis=1)
        pairs["was_active"] = pairs.apply(lambda x: x["item_scraped_count"] > 9, axis=1)
        return pairs

    @staticmethod
    def get_sum_scraped_currently_stopped(processed_pairs):
        currently_stopped = processed_pairs[processed_pairs["is_stopped"]]
        currently_stopped = currently_stopped.sort_values(["territory_uid", "updated_at"], ascending=False)

        # If the sum of scrapped document is 0 between two collects, that means that the collect has been inactive
        # in-between.
        currently_stopped["sum_scrapped"] = currently_stopped.groupby(["territory_uid"])[
            "item_scraped_count"].cumsum()

        return currently_stopped

    @staticmethod
    def filter_was_active(currently_stopped, last_active_day_treshold=12):
        """Extract collects that have been inactive for last_active_day or more.
        Collects that were never active are excluded.
        """
        # We only want to keep collects that have been active at some point.
        ever_active_collects = currently_stopped.groupby("territory_uid")["was_active"]\
            .agg(lambda x: any(x)).reset_index()
        ever_active_uids = ever_active_collects[ever_active_collects["was_active"]]["territory_uid"]

        # Combining Active collect, no collect inbetween, and interval between collects above treshold, we have the
        # candidate.
        stopped_collect = currently_stopped[(currently_stopped["territory_uid"].isin(ever_active_uids)) &
                                            (currently_stopped["sum_scrapped"] == 0) &
                                            (currently_stopped["interval"] >= timedelta(last_active_day_treshold))]

        # We keep only the furthest inactive collect.
        return stopped_collect.loc[stopped_collect.groupby(["territory_uid"])["interval"].idxmax()]

    @staticmethod
    def compute_timedelta(x):
        return x["updated_at_last"] - x["updated_at"]

    @staticmethod
    def is_collect_stopped(x):
        return x["item_scraped_count_last"] == 0

    @staticmethod
    def was_active(x):
        return x["item_scraped_count"] > 9
