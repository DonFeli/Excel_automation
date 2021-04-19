
class CollectFields:
    id = 'id'
    territory_uid = 'territory_uid'
    updated_at = 'updated_at'
    website = 'website'
    item_scraped_count = 'item_scraped_count'
    finish_reason = 'finish_reason'
    status = 'status'


class SheetParameters:
    """Define parameters for individual columns in the project's worksheets"""
    class UpdatedAt:
        column = 1
        name = CollectFields.updated_at
        range = "A1:A10000"
        width = 20

    class Website:
        column = 2
        name = CollectFields.website
        range = "B1:B10000"
        width = 40

    class TerritoryUid:
        column = 3
        name = CollectFields.territory_uid
        range = "C1:C10000"
        width = 15

    class Id:
        column = 4
        name = CollectFields.id
        range = 'D1:D10000'
        width = 8

    class ItemScrapedCount:
        column = 5
        name = CollectFields.item_scraped_count
        range = "E1:E10000"
        width = 22

    class FinishReason:
        column = 6
        name = CollectFields.finish_reason
        range = "F1:F10000"
        width = 13

    class Status:
        column = 7
        name = CollectFields.status
        range = "G1:G10000"
        width = 13

    class UpdatedAtLast:
        column = 8
        range = "H1:H10000"
        width = 20

    class ItemScrapedCountLast:
        column = 9
        range = "I1:I10000"
        width = 22
