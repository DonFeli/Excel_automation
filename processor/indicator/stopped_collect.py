from datetime import timedelta

from config.definitions import CollectFields


class StoppedCollectDetector:
    def __init__(self, collects):
        self.collect_data = self.filter_insufficient_collects(collects, 2)
        self.pairs = self.get_last_collect_by_territory(self.collect_data)
        self.processed_pairs = self.get_processed_pairs(self.pairs)

    @staticmethod
    def filter_insufficient_collects(collect_data, minimal_collect_count):
        """Exclude collects from territories where the total number of collect is less than minimal_collect_count."""
        return collect_data.groupby(
            CollectFields.territory_uid
        ).filter(lambda x: x[CollectFields.id].count() >= minimal_collect_count)

    @staticmethod
    def get_pairs_with_last_collect(collect_data):
        """Pair each collect with the last one on a given territory"""
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
        """Calculates the cumulative sum of items scraped for each territory where the last collect is stopped"""
        currently_stopped = processed_pairs[processed_pairs["is_stopped"]]
        currently_stopped = currently_stopped.sort_values(["territory_uid", "updated_at"], ascending=False)
        currently_stopped["sum_scrapped"] = currently_stopped.groupby(["territory_uid"])[
            "item_scraped_count"].cumsum()
        return currently_stopped

    @staticmethod
    def get_ever_active(currently_stopped):
        """Extracts territories with at least one active collect."""
        ever_active = currently_stopped.groupby("territory_uid")["was_active"].agg(lambda x: any(x)).reset_index()
        return ever_active

    @staticmethod
    def get_candidate(ever_active_collects, currently_stopped, last_active_day_treshold=12):
        ever_active_uids = ever_active_collects[ever_active_collects["was_active"]]["territory_uid"]

        # Combining Active collect, no collect inbetween, and interval between collects above treshold, we have the
        # candidate.
        stopped_collect = currently_stopped[(currently_stopped["territory_uid"].isin(ever_active_uids)) &
                                            (currently_stopped["sum_scrapped"] == 0) &
                                            (currently_stopped["interval"] >= timedelta(last_active_day_treshold))]

        # We keep only the furthest inactive collect.
        return stopped_collect.loc[stopped_collect.groupby(["territory_uid"])["interval"].idxmax()]