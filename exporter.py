from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image
import os

class ExcelReport:
	def __init__(self):
		"""Create a report in Excel, for reporting use cases."""
		self.wb = Workbook()

	def populate_sheet(self, data, ws_title, autofit=True):
		"""Create a new Excel sheet and populate it with data from a dataframe, passing in a specified string for the sheet title (limited to 30 characters)."""

		ws = self.wb.create_sheet(title=ws_title)
		if isinstance(data, pd.DataFrame):
			if len(ws_title) > 30:
				raise NotImplementedError("Sheet title cannot be longer than 30 characters. Title is currently {} characters long.".format(len(ws_title)))
			else:
				print('\tPopulating sheet...')
				for r in dataframe_to_rows(data, index=False, header=True):
					ws.append(r)

				if autofit:
					for col in ws.columns:
						length = len(col[0].value) + 3
						ws.column_dimensions[col[0].column].width = length

				print('\t...Done populating.')
		else:
			raise NotImplementedError("Input data must be a Pandas DataFrame.")

	def add_image(self, ws_title, img_path, anchor):
		"""Add an image to worksheet, specifying the anchor cell (for the top-left corner of the image)."""
		ws = self.wb[ws_title]
		ws.add_image(Image(img_path), anchor)
		print('\tAdded {img} to {sheet} at cell {cell}'.format(img=os.path.basename(img_path), sheet=ws_title, cell=anchor))

	def save_report(self, path):
		"""Save the workbook to specified filepath."""
		self.wb.remove_sheet(self.wb.active)
		self.wb.save(path)
		print('Done saving file.')