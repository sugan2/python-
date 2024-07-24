import os.path
from dataclasses import dataclass
import win32com.client as win32
from dataclasses import dataclass
from dmsService import Dms
import time
from tulipService import Bot
from tulipService.model.variableModel import VariableModel


@dataclass
class BotInputSchema:
    dmsCredentials: VariableModel = None
    domainName: VariableModel = None
    otc28GeneratedReport: VariableModel = None
    totalRowCount: VariableModel = None
    riskCategoryRowCount: VariableModel = None
    creditLimitRowCount: VariableModel = None
    combinedRowCount: VariableModel = None
    sheetNumber: VariableModel = None
    headerRow: VariableModel = None
    valueRow: VariableModel = None
    fileName: VariableModel = None


@dataclass
class BotOutputSchema:
    def __init__(self):
        self.Oct28finalReport = VariableModel()


class BotLogic(Bot):
    def __init__(self) -> None:
        super().__init__()
        try:
            self.outputs = BotOutputSchema()
            self.input = self.bot_input.get_proposedBotInputs(BotInputs=BotInputSchema)
            self.dms_identity = self.bot_input.get_identity(self.input.dmsCredentials.value)
            self._dms = Dms(beekeeper_url=self.input.domainName.value,
                            user_name=self.dms_identity.credential.basicAuth.username,
                            password=self.dms_identity.credential.basicAuth.password,
                            logger=self.log)
            self.full_path = None
        except Exception as error:
            self.log.error(f"Error initializing BotLogic: {error}")

    def paste_counts_to_excel(self, destination_file, sheetNumberkey, headerRow, value_row_number, total_row_count,
                              riskcategory,
                              creditlimit,
                              combined_value):

        excel = win32.Dispatch('Excel.Application')
        workbook = excel.Workbooks.Open(destination_file)
        sheet = workbook.Sheets(sheetNumberkey)

        headers = ["Total Rows Count", "Risk category", "Credit limit", "Combined Value"]
        values = [total_row_count, riskcategory, creditlimit, combined_value]

        for col, header in enumerate(headers, start=1):
            sheet.Cells(headerRow, col).Value = header
            sheet.Cells(value_row_number, col).Value = values[col - 1]

        workbook.Save()
        workbook.Close()
        excel.Quit()

    def main(self):
        try:

            self.log.info("DMS Server is connected successfully.")
            time.sleep(3)

        except Exception as E:
            self.log.error(f"{E}")

        destination_file = self.input.otc28GeneratedReport.value
        download_destination_file = self._dms.download_file_dms(file_signature=destination_file,
                                                                save_directory=self.working_dir)

        sheetNumberkey = self.input.sheetNumber.value
        header_row_number = self.input.headerRow.value
        value_row_number = self.input.valueRow.value
        total_row_count = self.input.totalRowCount.value
        riskcategory = self.input.riskCategoryRowCount.value
        creditlimit = self.input.creditLimitRowCount.value
        combined_value = self.input.combinedRowCount.value
        self.log.info("Took all values.")

        try:
            self.paste_counts_to_excel(download_destination_file, sheetNumberkey, header_row_number,
                                       value_row_number,
                                       total_row_count, riskcategory, creditlimit, combined_value)
            self.log.info("paste done successfully.")
        except Exception as e:
            self.log.error(f"error in pasting in excel")

        try:

            finalMOSt = self._dms.upload_file_to_dms(download_destination_file)
            self.log.info("uploaded to dms.")
            # screenshotfilesignature=self.input.Oct28finalReport.value
            self.bot_output.add_variable(key=self.input.fileName.value, val=finalMOSt)
            #self.outputs.Oct28finalReport.value = finalMOSt
            self.bot_output.success()
        except Exception as e:
            self.log.error(f"error in uploading final report to dms")


