from argparse import ArgumentParser
from config import logger
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.formula.translate import Translator
import pandas as pd

from processor.extractor.extractor import CollectOperationExtractor
from processor.transformer.collect_transformer import CollectTransformer
from processor.indicator.stopped_collect import StoppedCollectDetector
from config.definitions import SheetParameters

pd.set_option("display.max_columns", 500)
pd.set_option("display.width", 0)


def parse_arguments():
    """Parse argument from console"""
    parser = ArgumentParser()
    parser.add_argument('--gsheet', '-g', type=str, default='Excel automation project', help='Name of excel sheet.')
    return parser.parse_args()


def print_rows(sheet):
    for row in sheet.iter_rows(values_only=True):
        print(row)


def main(args):
    logger.info('-------------- Loading and transforming data --------------------')
    collect_operations = CollectOperationExtractor()
    collect_operations.retrieve_collects()
    structured_collects = CollectTransformer().load_and_transform(collect_operations.collect_data)

    # INITIAL COLLECTS

    logger.info('-------------- Sheet 1 : Initial Collects --------------------')
    workbook = Workbook()
    sheet1 = workbook.active
    sheet1.title = "Initial Collects"

    for row in dataframe_to_rows(structured_collects, index=False, header=True):
        sheet1.append(row)

    # SUFFICIENT COLLECTS

    logger.info('-------------- Sheet 2 : Sufficient Collects --------------------')
    sufficient_collects = StoppedCollectDetector.filter_insufficient_collects(structured_collects, 2)
    sufficient_collects_sheet = workbook.create_sheet("Sufficient Collects")
    for row in dataframe_to_rows(sufficient_collects, index=False, header=True):
        sufficient_collects_sheet.append(row)

    # PAIRS WITH LAST COLLECTS

    logger.info('-------------- Sheet 3 : Pairs with Last Collect --------------------')
    pairs_with_last_collect = StoppedCollectDetector.get_last_collect_by_territory(sufficient_collects)
    pairs_with_last_collect_sheet = workbook.create_sheet("Pairs with Last Collect")
    for row in dataframe_to_rows(pairs_with_last_collect, index=False, header=True):
        pairs_with_last_collect_sheet.append(row)

    # PROCESSED PAIRS WITH EXCEL

    logger.info('-------------- Sheet 4 : Processed Pairs with Excel --------------------')
    processed_pairs_excel_sheet = workbook.copy_worksheet(pairs_with_last_collect_sheet)
    processed_pairs_excel_sheet.title = 'Processed Pairs Excel'

    processed_pairs_excel_sheet['J1'].value = "interval"
    processed_pairs_excel_sheet['J2'] = "=(H2-A2)"

    processed_pairs_excel_sheet['K1'].value = "is_stopped"
    processed_pairs_excel_sheet['K2'] = '=IF(I2=0, "TRUE", "FALSE")'

    processed_pairs_excel_sheet['L1'].value = "was_active"
    processed_pairs_excel_sheet['L2'] = '=IF(E2>9, "TRUE", "FALSE")'

    for i in range(2, processed_pairs_excel_sheet.max_row):
        processed_pairs_excel_sheet[f'J{i}'] = Translator("=(H2-A2)", origin="J2").translate_formula(f'J{i}')
        processed_pairs_excel_sheet[f'K{i}'] = Translator('=IF(I2=0, "TRUE", "FALSE")', origin="K2").translate_formula(
            f'K{i}')
        processed_pairs_excel_sheet[f'L{i}'] = Translator('=IF(E2>9, "TRUE", "FALSE")', origin="L2").translate_formula(
            f'L{i}')

    # PROCESSED PAIRS WITH PYTHON

    logger.info('-------------- Sheet 5 : Processed Pairs with Python --------------------')

    processed_pairs = StoppedCollectDetector.get_processed_pairs(pairs_with_last_collect)
    processed_pairs['interval'] = processed_pairs['interval'].dt.days
    processed_pairs.sort_values(["territory_uid", "updated_at"], inplace=True, ascending=False)

    processed_pairs_python_sheet = workbook.create_sheet("Processed pairs Python")

    for row in dataframe_to_rows(processed_pairs, index=False, header=True):
        processed_pairs_python_sheet.append(row)

    # SUM SCRAPED ON CURRENTLY STOPPED COLLECTS

    logger.info('-------------- Sheet 6 : Currently Stopped --------------------')

    currently_stopped_sum_scraped = StoppedCollectDetector.get_sum_scraped_currently_stopped(processed_pairs)

    currently_stopped_sum_scraped_sheet = workbook.create_sheet('Sum Scraped on Currently Stopped Collects')

    for row in dataframe_to_rows(currently_stopped_sum_scraped, index=False, header=True):
        currently_stopped_sum_scraped_sheet.append(row)

    # EVER ACTIVE

    logger.info('-------------- Sheet 7 : Ever Active --------------------')

    ever_active_sheet = workbook.copy_worksheet(currently_stopped_sum_scraped_sheet)
    ever_active_sheet.title = 'Ever Active'

    # Récupère les territory_uids où la collecte était active (was_active = True)
    ever_active_rows = []
    for row in ever_active_sheet.iter_rows(min_row=2, max_row=ever_active_sheet.max_row, min_col=12, max_col=12):
        for cell in row:
            if cell.value == True:
                ever_active_rows.append(row[0].row)

    ever_active_uids = []
    for active_row in ever_active_rows:
        ever_active_uids.append(ever_active_sheet[f'C{active_row}'].value)

    # Supprime les collectes de territoires qui n'ont jamais été actifs
    for cell in ever_active_sheet['C']:
        if cell.row != 1:
            if cell.value not in ever_active_uids:
                ever_active_sheet.delete_rows(cell.row)

    # SCRAPPING DEAD

    logger.info('-------------- Sheet 8 : Scrapping Kaput --------------------')

    scrapping_kaput_sheet = workbook.copy_worksheet(ever_active_sheet)
    scrapping_kaput_sheet.title = 'Scrapping Kaput'

    # Supprime les collectes de territoires qui n'ont jamais été actifs
    for cell in scrapping_kaput_sheet['M']:
        if cell.row != 1:
            if cell.value != 0:
                scrapping_kaput_sheet.delete_rows(cell.row)

    # INTERVAL EXCEEDED 

    logger.info('-------------- Sheet 9 : Interval exceeded --------------------')

    interval_exceeded_sheet = workbook.copy_worksheet(scrapping_kaput_sheet)
    interval_exceeded_sheet.title = 'Scrapping Kaput'

    # # Supprime les lignes où l'interval est inférieur à 12 jours
    for cell in interval_exceeded_sheet['J']:
        if cell.row != 1:
            if cell.value < 12:
                interval_exceeded_sheet.delete_rows(cell.row)

    # SUMMARY STATISTICS

    logger.info('-------------- Sheet 10 : Interval exceeded --------------------')

    summary_statistics_sheet = workbook.create_sheet("Summary Statistics")
    # summary_statistics_sheet['A1'] = '=COUNTA(N3:N236)'
    # summary_statistics_sheet['A2'] = '=UNIQUE(C2:C1016)'

    # logger.info('-------------- Saving Excel file ---------------------------------')
    for sheet in workbook._sheets:
        sheet.column_dimensions[
            get_column_letter(SheetParameters.UpdatedAt.column)].width = SheetParameters.UpdatedAt.width
        sheet.column_dimensions[get_column_letter(SheetParameters.Website.column)].width = SheetParameters.Website.width
        sheet.column_dimensions[
            get_column_letter(SheetParameters.TerritoryUid.column)].width = SheetParameters.TerritoryUid.width
        sheet.column_dimensions[get_column_letter(SheetParameters.Id.column)].width = SheetParameters.Id.width
        sheet.column_dimensions[
            get_column_letter(SheetParameters.ItemScrapedCount.column)].width = SheetParameters.ItemScrapedCount.width
        sheet.column_dimensions[
            get_column_letter(SheetParameters.FinishReason.column)].width = SheetParameters.FinishReason.width
        sheet.column_dimensions[get_column_letter(SheetParameters.Status.column)].width = SheetParameters.Status.width
        sheet.column_dimensions[
            get_column_letter(SheetParameters.UpdatedAtLast.column)].width = SheetParameters.UpdatedAtLast.width
        sheet.column_dimensions[get_column_letter(
            SheetParameters.ItemScrapedCountLast.column)].width = SheetParameters.ItemScrapedCountLast.width
        sheet.auto_filter.ref = sheet.dimensions

    workbook.save(filename=f'{args.gsheet}.xlsx')


if __name__ == '__main__':
    args = parse_arguments()
    main(args)
    # # Supprime les lignes où les collectes ne sont pas arrêtées (is_stopped = False)
    # for cell in filtered_pairs_sheet['K']:
    #     if cell.value == False:
    #         filtered_pairs_sheet.delete_rows(cell.row)
