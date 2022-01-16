from MainWindow import Ui_MainWindow

from PyQt5 import QtWidgets as qtw
from PyQt5 import QtCore as qtc
from PyQt5 import QtGui as qtg

from copy import copy, deepcopy
import pandas as pd
import json
import os
from datetime import date


class MainWindow(qtw.QMainWindow):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.actionSave_As.setEnabled(False)
        self.ui.actionSave.setEnabled(False)
        self.ui.tabWidget.setTabEnabled(3, False)
        self.ui.tabWidget.setTabEnabled(2, False)
        self.ui.tabWidget.setTabEnabled(1, False)
        self.ui.tabWidget.setTabEnabled(0, False)

        # Push buttons
        ##### Stock Sheet #####
        # self.ui.edit_button.setEnabled(False)
        self.ui.clear_button.setEnabled(False)
        self.ui.delete_entry.setEnabled(False)
        self.ui.reset_find_preview.setEnabled(False)
        self.ui.confirm_edit.setEnabled(False)
        self.ui.confirm_edit.clicked.connect(
            lambda: self.confirm_edit(tab="stock"))
        self.ui.Add_stock.clicked.connect(self.add_stock)
        self.ui.find_button.clicked.connect(
            lambda: self.find_item_in_stock_sheet(sheet="stock_sheet"))
        self.ui.delete_entry.clicked.connect(
            lambda: self.delete_entry_from_tabel(tab="stock"))
        self.ui.clear_button.clicked.connect(
            lambda: self.clear_tabel(self.ui.current_stock_find_entry))
        self.ui.clear_button.clicked.connect(
            lambda: self.clear_tabel(self.ui.find_stock_preview, True))
        self.ui.clear_button.clicked.connect(self.stock_tab_button_set_False)
        self.ui.reset_find_preview.clicked.connect(
            lambda: self.reset_preview(tab="stock"))
        self.ui.current_stock_find_entry.keyPressEvent = self.find_enter_event
        self.ui.find_stock_preview.keyPressEvent = self.preview_enter_event
        self.ui.scrollArea_2.mousePressEvent = self.reset_selected

        ##### Issue Sheet #####
        self.ui.clear_button_2.setEnabled(False)
        self.ui.delete_entry_2.setEnabled(False)
        self.ui.edit_button_2.setEnabled(False)
        self.ui.Reset_issued_preview.setEnabled(False)
        self.ui.approval_issue.setEnabled(False)
        self.ui.reset_issue_items.setEnabled(False)
        self.ui.Delete_issue_items.setEnabled(False)
        self.ui.add_sku.setEnabled(False)
        self.ui.Weight_to_add.setEnabled(False)
        self.ui.find_button_2.clicked.connect(
            lambda: self.find_item_in_stock_sheet(sheet="issue_sheet"))
        self.ui.clear_button_2.clicked.connect(
            lambda: self.clear_tabel(self.ui.issued_stock_find_entry))
        self.ui.clear_button_2.clicked.connect(
            lambda: self.clear_tabel(self.ui.issue_stock_preview, True))
        self.ui.clear_button_2.clicked.connect(self.issue_tab_button_set_false)
        self.ui.edit_button_2.clicked.connect(
            lambda: self.confirm_edit(tab="issue"))
        self.ui.Reset_issued_preview.clicked.connect(
            lambda: self.reset_preview(tab="issue"))
        self.ui.delete_entry_2.clicked.connect(
            lambda: self.delete_entry_from_tabel(tab="issue"))
        self.ui.find_sku.clicked.connect(self.find_sku)
        self.ui.clear_issue_stock.clicked.connect(
            lambda: self.clear_tabel(self.ui.sku_find, True))
        self.ui.clear_issue_stock.clicked.connect(
            lambda: self.ui.add_sku.setEnabled(False))
        self.ui.clear_issue_stock.clicked.connect(
            lambda: self.ui.enter_sku_code.clear())
        self.ui.Delete_issue_items.clicked.connect(self.delete_item_from_issue_table)
        self.ui.reset_issue_items.clicked.connect(self.reset_issue_items)
        # self.ui.clear_issue_stock.clicked.connect(
        #     lambda: self.ui.Weight_to_add.clear())
        self.ui.add_sku.clicked.connect(self.add_sku)
        self.ui.dateEdit_2.setDate(
            qtc.QDate(date.today().year, date.today().month, date.today().day))
        self.ui.scrollArea.mousePressEvent = self.reset_selected_issue

        # tabel
        # self.ui.find_stock_preview.clicked.connect(lambda : self.ui.delete_entry.setEnabled(True))

        # Variabels
        self.saved = True
        self.deleted = []
        self.original = []
        self.original_issue = []
        self.sheet_index_add_sku = None
        self.issue_sheet_skus = []
        self.issue_sheet_items = []
        # self.original_prev = []

        # Action buttons
        self.ui.actionOpen_Project.triggered.connect(self.open_project)
        self.ui.actionExit_2.triggered.connect(self.exit)
        self.ui.actionNew_Project.triggered.connect(self.New_project)
        self.ui.actionSave.triggered.connect(
            lambda: self.save_project(save_path=self.project_path))
        self.ui.actionSave_As.triggered.connect(lambda: self.save_project(qtw.QFileDialog.getExistingDirectory(
            self, "Select Directory")))

    def red_palette(self):
        palette = qtg.QPalette()
        brush = qtg.QBrush(qtg.QColor(0, 0, 0))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active, qtg.QPalette.WindowText, brush)
        brush = qtg.QBrush(qtg.QColor(239, 41, 41))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active, qtg.QPalette.Button, brush)
        brush = qtg.QBrush(qtg.QColor(255, 147, 147))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active, qtg.QPalette.Light, brush)
        brush = qtg.QBrush(qtg.QColor(247, 94, 94))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active, qtg.QPalette.Midlight, brush)
        brush = qtg.QBrush(qtg.QColor(119, 20, 20))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active, qtg.QPalette.Dark, brush)
        brush = qtg.QBrush(qtg.QColor(159, 27, 27))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active, qtg.QPalette.Mid, brush)
        brush = qtg.QBrush(qtg.QColor(0, 0, 0))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active, qtg.QPalette.Text, brush)
        brush = qtg.QBrush(qtg.QColor(255, 255, 255))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active, qtg.QPalette.BrightText, brush)
        brush = qtg.QBrush(qtg.QColor(0, 0, 0))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active, qtg.QPalette.ButtonText, brush)
        brush = qtg.QBrush(qtg.QColor(255, 255, 255))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active, qtg.QPalette.Base, brush)
        brush = qtg.QBrush(qtg.QColor(239, 41, 41))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active, qtg.QPalette.Window, brush)
        brush = qtg.QBrush(qtg.QColor(0, 0, 0))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active, qtg.QPalette.Shadow, brush)
        brush = qtg.QBrush(qtg.QColor(247, 148, 148))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active,
                         qtg.QPalette.AlternateBase, brush)
        brush = qtg.QBrush(qtg.QColor(255, 255, 220))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active, qtg.QPalette.ToolTipBase, brush)
        brush = qtg.QBrush(qtg.QColor(0, 0, 0))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active, qtg.QPalette.ToolTipText, brush)
        brush = qtg.QBrush(qtg.QColor(0, 0, 0, 128))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Active,
                         qtg.QPalette.PlaceholderText, brush)
        brush = qtg.QBrush(qtg.QColor(0, 0, 0))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive, qtg.QPalette.WindowText, brush)
        brush = qtg.QBrush(qtg.QColor(239, 41, 41))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive, qtg.QPalette.Button, brush)
        brush = qtg.QBrush(qtg.QColor(255, 147, 147))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive, qtg.QPalette.Light, brush)
        brush = qtg.QBrush(qtg.QColor(247, 94, 94))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive, qtg.QPalette.Midlight, brush)
        brush = qtg.QBrush(qtg.QColor(119, 20, 20))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive, qtg.QPalette.Dark, brush)
        brush = qtg.QBrush(qtg.QColor(159, 27, 27))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive, qtg.QPalette.Mid, brush)
        brush = qtg.QBrush(qtg.QColor(0, 0, 0))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive, qtg.QPalette.Text, brush)
        brush = qtg.QBrush(qtg.QColor(255, 255, 255))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive, qtg.QPalette.BrightText, brush)
        brush = qtg.QBrush(qtg.QColor(0, 0, 0))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive, qtg.QPalette.ButtonText, brush)
        brush = qtg.QBrush(qtg.QColor(255, 255, 255))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive, qtg.QPalette.Base, brush)
        brush = qtg.QBrush(qtg.QColor(239, 41, 41))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive, qtg.QPalette.Window, brush)
        brush = qtg.QBrush(qtg.QColor(0, 0, 0))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive, qtg.QPalette.Shadow, brush)
        brush = qtg.QBrush(qtg.QColor(247, 148, 148))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive,
                         qtg.QPalette.AlternateBase, brush)
        brush = qtg.QBrush(qtg.QColor(255, 255, 220))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive,
                         qtg.QPalette.ToolTipBase, brush)
        brush = qtg.QBrush(qtg.QColor(0, 0, 0))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive,
                         qtg.QPalette.ToolTipText, brush)
        brush = qtg.QBrush(qtg.QColor(0, 0, 0, 128))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Inactive,
                         qtg.QPalette.PlaceholderText, brush)
        brush = qtg.QBrush(qtg.QColor(119, 20, 20))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled, qtg.QPalette.WindowText, brush)
        brush = qtg.QBrush(qtg.QColor(239, 41, 41))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled, qtg.QPalette.Button, brush)
        brush = qtg.QBrush(qtg.QColor(255, 147, 147))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled, qtg.QPalette.Light, brush)
        brush = qtg.QBrush(qtg.QColor(247, 94, 94))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled, qtg.QPalette.Midlight, brush)
        brush = qtg.QBrush(qtg.QColor(119, 20, 20))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled, qtg.QPalette.Dark, brush)
        brush = qtg.QBrush(qtg.QColor(159, 27, 27))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled, qtg.QPalette.Mid, brush)
        brush = qtg.QBrush(qtg.QColor(119, 20, 20))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled, qtg.QPalette.Text, brush)
        brush = qtg.QBrush(qtg.QColor(255, 255, 255))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled, qtg.QPalette.BrightText, brush)
        brush = qtg.QBrush(qtg.QColor(119, 20, 20))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled, qtg.QPalette.ButtonText, brush)
        brush = qtg.QBrush(qtg.QColor(239, 41, 41))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled, qtg.QPalette.Base, brush)
        brush = qtg.QBrush(qtg.QColor(239, 41, 41))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled, qtg.QPalette.Window, brush)
        brush = qtg.QBrush(qtg.QColor(0, 0, 0))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled, qtg.QPalette.Shadow, brush)
        brush = qtg.QBrush(qtg.QColor(239, 41, 41))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled,
                         qtg.QPalette.AlternateBase, brush)
        brush = qtg.QBrush(qtg.QColor(255, 255, 220))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled,
                         qtg.QPalette.ToolTipBase, brush)
        brush = qtg.QBrush(qtg.QColor(0, 0, 0))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled,
                         qtg.QPalette.ToolTipText, brush)
        brush = qtg.QBrush(qtg.QColor(0, 0, 0, 128))
        brush.setStyle(qtc.Qt.SolidPattern)
        palette.setBrush(qtg.QPalette.Disabled,
                         qtg.QPalette.PlaceholderText, brush)

        return palette

    def add_sku(self):
        # Validate if approval number is unique
        approval_no = self.ui.approval_no_2.value()
        self.ui.approval_no_2.setPalette(self.style().standardPalette())
        self.ui.broker_name.setPalette(self.style().standardPalette())
        self.ui.party_name.setPalette(self.style().standardPalette())
        self.ui.Weight_to_add.setPalette(self.style().standardPalette())
        valid = True
        add = True
        for row in self.issue_sheet:
            if approval_no == row["APPROVAL NO."]:
                valid = False
                break
        if valid:
            for row in self.issue_return_sheet:
                if approval_no == row["APPROVAL NO."]:
                    valid = False
                    break
        if self.ui.broker_name.text() == '':
            self.ui.broker_name.setPalette(self.red_palette())
            self.ui.statusbar.showMessage("Enter Broker Name")
            add = False
        if self.ui.party_name.text() == '':
            self.ui.party_name.setPalette(self.red_palette())
            self.ui.statusbar.showMessage("Enter Party Name")
            add = False
        if self.ui.Weight_to_add.value() == 0:
            self.ui.Weight_to_add.setPalette(self.red_palette())
            self.ui.statusbar.showMessage("Weight can not be 0")
            add = False
        if not valid:
            self.ui.statusbar.showMessage("Approval No. Invalid")
            self.ui.approval_no_2.setPalette(self.red_palette())
            add = False
        if add:
            # fix broker name and party name
            self.ui.broker_name.setReadOnly(True)
            self.ui.party_name.setReadOnly(True)
            self.ui.dateEdit_2.setReadOnly(True)
            self.ui.approval_no_2.setReadOnly(True)

            item_to_add = {"stock index": self.sheet_index_add_sku}
            item_to_add = deepcopy(self.stock_sheet[self.sheet_index_add_sku])
            # check if sku aleady in issue_sheet_skus
            if item_to_add["SKU"] in self.issue_sheet_skus:
                self.ui.sku_find.setPalette(self.red_palette())
                self.ui.statusbar.showMessage("SKU already in approval")
                return
            self.issue_sheet_skus.append(copy(item_to_add["SKU"]))
            self.stock_sheet[self.sheet_index_add_sku]["Weight in Hand"] -= self.ui.Weight_to_add.value()
            self.stock_sheet[self.sheet_index_add_sku]["Issue Qty"] += self.ui.Weight_to_add.value()
            for key in ["Weight in Hand",
                        "Short/Excess Received",
                        "Issue Qty",
                        "Sold Qty",
                        "# Code",
                        "Remarks",
                        "Date"]:
                del(item_to_add[key])
            item_to_add["Approval Date"] = self.ui.dateEdit_2.date().toString()
            item_to_add["APPROVAL NO."] = self.ui.approval_no_2.value()
            item_to_add["BROKER NAME"] = self.ui.broker_name.text()
            item_to_add["PARTY NAME"] = self.ui.party_name.text()
            item_to_add["Weight issued"] = self.ui.Weight_to_add.value()
            item_dict = {}
            for key in self.schema["issue_sheet"]["columns"]:
                if item_to_add.get(key):
                    item_dict[key] = item_to_add[key]
                else:
                    item_dict[key] = ''
            self.ui.issue_items.setEnabled(True)
            self.issue_sheet_items.append(deepcopy(item_dict))
            self.load_data_in_tabel(
                self.ui.issue_items, values=item_dict,
                rowcount=self.ui.issue_items.rowCount()+1,
                columns=list(item_dict),
                affected_row=self.ui.issue_items.rowCount(), read_only=["BROKER NAME",
                                                                        "PARTY NAME", "Approval Date",
                                                                        "APPROVAL NO.", "Weight issued"])

            self.ui.approval_issue.setEnabled(True)
            self.ui.Delete_issue_items.setEnabled(True)
            self.ui.reset_issue_items.setEnabled(True)
            self.ui.clear_issue_stock.click()

    def issue_approval(self):
        self.ui.approval_issue.setEnabled(False)
        self.ui.Delete_issue_items.setEnabled(False)
        self.ui.reset_issue_items.setEnabled(False)
        self.clear_tabel(self.ui.issue_items, disable=True, row_count=0)

        # load values from table to issue_sheet_items
        for row in range(len(self.issue_sheet_items)):
            for col, index in self.issue_sheet_items[row]:
                if not bool(self.issue_sheet_items[row][col]):
                    self.issue_sheet_items[row][col] = self.ui.issue_items.item(row, index).text()
        
        doc = docx.Document


    def delete_item_from_issue_table(self):
        if len(self.ui.issue_items.selectedIndexes()) == self.ui.issue_items.columnCount() and self.ui.issue_items.selectedIndexes().count(self.ui.issue_items.selectedIndexes()[0]) == len(self.ui.issue_items.selectedIndexes()):
            # save item
            sheet_index = self.issue_sheet_items[self.ui.issue_items.selectedIndexes()[0].row()]["stock index"]
            self.stock_sheet[sheet_index]["Weight in Hand"]+=self.issue_sheet_items[self.ui.issue_items.selectedIndexes()[0].row()]["Weight issued"]
            self.stock_sheet[sheet_index]["Issue Qty"]-=self.issue_sheet_items[self.ui.issue_items.selectedIndexes()[0].row()]["Weight issued"]
            del(self.issue_sheet_items[self.ui.issue_items.selectedIndexes()[0].row()])
            self.ui.issue_items.removeRow(
                self.ui.issue_items.selectedIndexes()[0].row())
        else:
            indexes = self.ui.issue_items.selectedIndexes()
            for index in indexes:
                self.ui.issue_items.setItem(
                    index.row(), index.column(), qtw.QTableWidgetItem(str('')))
                self.issue_sheet[int(self.ui.issue_items.item(index.row(
                ), 0).text())-1][self.ui.issue_items.horizontalHeaderItem(index.column()).text()] = ''
            # self.load_data_in_tabel(self.ui.current_stock, self.issue_sheet, len(
            #     self.issue_sheet), self.schema["stock_sheet"]["columns"])
    
    def reset_issue_items(self):
        for index in range(self.ui.issue_items.rowCount()):
            sheet_index = self.issue_sheet_items[index]["stock index"]
            self.stock_sheet[sheet_index]["Weight in Hand"]+=self.issue_sheet_items[index]["Weight issued"]
            self.stock_sheet[sheet_index]["Issue Qty"]-=self.issue_sheet_items[index]["Weight issued"]
            del(self.issue_sheet_items[index])
            self.ui.issue_items.removeRow(
                index)

    def find_sku(self):
        self.clear_tabel(self.ui.sku_find, disable=True)
        self.ui.enter_sku_code.setPalette(self.style().standardPalette())
        self.ui.add_sku.setEnabled(False)
        sku_to_find = self.ui.enter_sku_code.text()
        if not bool(sku_to_find):
            return
        sku_found = False
        count = 0
        for dicti in self.stock_sheet:
            if dicti["SKU"] == sku_to_find:
                sku_found = dicti
                self.sheet_index_add_sku = count
                break
            count += 1
        if sku_found != False:
            self.load_data_in_tabel(self.ui.sku_find, values=[
                                    sku_found], rowcount=1, columns=self.schema["stock_sheet"]["columns"])
            self.ui.sku_find.setEnabled(True)
            self.ui.add_sku.setEnabled(True)
            self.ui.Weight_to_add.setEnabled(True)
            self.ui.Weight_to_add.setValue(float(sku_found["Weight in Hand"]))
            self.ui.Weight_to_add.setMaximum(
                float(sku_found["Weight in Hand"]))
        else:
            self.ui.enter_sku_code.setPalette(self.red_palette())

    def add_stock(self):
        entry_dict = {}
        for j in range(self.ui.current_entry_add_stock.columnCount()):
            try:
                entry_dict[self.ui.current_entry_add_stock.horizontalHeaderItem(
                    j).text()] = self.ui.current_entry_add_stock.item(0, j).text()
            except AttributeError:
                entry_dict[self.ui.current_entry_add_stock.horizontalHeaderItem(
                    j).text()] = self.ui.current_entry_add_stock.item(0, j)
        self.stock_sheet.append(entry_dict)
        # self.ui.current_entry_add_stock.clear()
        self.load_data_in_tabel(self.ui.current_stock, entry_dict, len(
            self.stock_sheet), list(entry_dict), len(self.stock_sheet)-1)
        tabel_columns = []
        for i in range(self.ui.current_entry_add_stock.columnCount()):
            tabel_columns.append(
                self.ui.current_entry_add_stock.horizontalHeaderItem(i).text())
        self.ui.current_entry_add_stock.clear()
        self.ui.current_entry_add_stock.setHorizontalHeaderLabels(
            tabel_columns)

    def confirm_edit(self, tab="stock"):
        if tab == "stock":
            for i in range(self.ui.find_stock_preview.rowCount()):
                for j in range(1, self.ui.find_stock_preview.columnCount()):

                    item = self.ui.find_stock_preview.item(i, j)

                    self.stock_sheet[int(self.ui.find_stock_preview.item(i, 0).text(
                    ))-1][self.ui.find_stock_preview.horizontalHeaderItem(j).text()] = item.text()
                    self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
                        self.stock_sheet), self.schema["stock_sheet"]["columns"])

        elif tab == "issue":
            for i in range(self.ui.issue_stock_preview.rowCount()):
                for j in range(1, self.ui.issue_stock_preview.columnCount()):

                    item = self.ui.issue_stock_preview.item(i, j)

                    self.issue_sheet[int(self.ui.issue_stock_preview.item(i, 0).text(
                    ))-1][self.ui.issue_stock_preview.horizontalHeaderItem(j).text()] = item.text()

                    ##### Add recalculate stock_sheet here ####

    def reset_selected(self, event):
        self.ui.current_stock_find_entry.clearSelection()
        self.ui.find_stock_preview.clearSelection()
        self.ui.current_entry_add_stock.clearSelection()
        self.ui.current_stock.clearSelection()

    def reset_selected_issue(self, event):
        self.ui.issue_stock_preview.clearSelection()
        self.ui.issued_stock_find_entry.clearSelection()
        self.ui.sku_find.clearSelection()
        self.ui.issue_items.clearSelection()
        if not self.ui.sku_find.isEnabled():
            self.ui.enter_sku_code.clear()
            self.ui.Weight_to_add.clear()
        if not bool(self.ui.enter_sku_code.text()):
            # self.ui.approval_no_2.palette()
            self.ui.enter_sku_code.setPalette(self.style().standardPalette())

    def issue_find_enter_event(self, event):
        # Enter key pressed
        if event.key() == qtc.Qt.Key_Return or event.key() == qtc.Qt.Key_Enter:
            if len(self.ui.issued_stock_find_entry.selectedIndexes()) == 1:
                item = self.ui.issued_stock_find_entry.itemFromIndex(
                    self.ui.issued_stock_find_entry.selectedIndexes()[0])
                if item is None:
                    self.ui.issued_stock_find_entry.edit(
                        self.ui.issued_stock_find_entry.selectedIndexes()[0])
                else:
                    self.find_item_in_stock_sheet(sheet="issue_sheet")
        qtw.QTableWidget.keyPressEvent(self.ui.issued_stock_find_entry, event)

    def issue_preview_enter_event(self, event):
        # Enter key pressed
        if event.key() == qtc.Qt.Key_Return or event.key() == qtc.Qt.Key_Enter:
            index = self.ui.issue_stock_preview.selectedIndexes()
            if len(index) == 1:
                index = index[0]
                item = self.ui.issue_stock_preview.itemFromIndex(index)

                self.stock_sheet[int(self.ui.issue_stock_preview.item(index.row(
                ), 0).text())-1][self.ui.issue_stock_preview.horizontalHeaderItem(index.column()).text()] = item.text()
                self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
                    self.stock_sheet), self.schema["stock_sheet"]["columns"])
        if event.key() == qtc.Qt.Key_Delete:
            self.delete_entry_from_tabel(tab="issue")

        qtw.QTableWidget.keyPressEvent(self.ui.issue_stock_preview, event)

    def find_enter_event(self, event):
        # Enter key pressed
        if event.key() == qtc.Qt.Key_Return or event.key() == qtc.Qt.Key_Enter:
            if len(self.ui.current_stock_find_entry.selectedIndexes()) == 1:
                item = self.ui.current_stock_find_entry.itemFromIndex(
                    self.ui.current_stock_find_entry.selectedIndexes()[0])
                if item is None:
                    self.ui.current_stock_find_entry.edit(
                        self.ui.current_stock_find_entry.selectedIndexes()[0])
                else:
                    self.find_item_in_stock_sheet(sheet="stock_sheet")
        qtw.QTableWidget.keyPressEvent(self.ui.current_stock_find_entry, event)

    def preview_enter_event(self, event):
        # Enter key pressed
        if event.key() == qtc.Qt.Key_Return or event.key() == qtc.Qt.Key_Enter:
            index = self.ui.find_stock_preview.selectedIndexes()
            if len(index) == 1:
                index = index[0]
                item = self.ui.find_stock_preview.itemFromIndex(index)

                self.stock_sheet[int(self.ui.find_stock_preview.item(index.row(
                ), 0).text())-1][self.ui.find_stock_preview.horizontalHeaderItem(index.column()).text()] = item.text()
                self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
                    self.stock_sheet), self.schema["stock_sheet"]["columns"])
        if event.key() == qtc.Qt.Key_Delete:
            self.delete_entry_from_tabel(tab="stock")

        qtw.QTableWidget.keyPressEvent(self.ui.find_stock_preview, event)

    def reset_preview(self, tab="stock"):
        if tab == "stock":
            for dicti in self.original:

                index = copy(int(dicti["sheet_index"]))
                temp_dict = deepcopy(dicti)
                del(temp_dict["sheet_index"])
                self.stock_sheet[index-1] = deepcopy(temp_dict)
            self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
                self.stock_sheet), self.schema["stock_sheet"]["columns"])
            self.load_data_in_tabel(self.ui.find_stock_preview, self.original, len(
                self.original), list(self.original[0]))

        elif tab == "issue":
            for dicti in self.original_issue:

                index = copy(int(dicti["sheet_index"]))
                temp_dict = deepcopy(dicti)
                del(temp_dict["sheet_index"])
                self.issue_sheet[index-1] = deepcopy(temp_dict)
            # self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
            #     self.stock_sheet), self.schema["stock_sheet"]["columns"])
            self.load_data_in_tabel(self.ui.issue_stock_preview, self.original_issue, len(
                self.original_issue), list(self.original_issue[0]))

    def clear_tabel(self, tabel, disable=False, row_count=1):

        assert(isinstance(tabel, qtw.QTableWidget))
        tabel_columns = []
        for i in range(tabel.columnCount()):
            tabel_columns.append(tabel.horizontalHeaderItem(i).text())
        tabel.clear()
        tabel.setHorizontalHeaderLabels(tabel_columns)
        tabel.setRowCount(row_count)
        if disable == True:
            tabel.setEnabled(False)

    def stock_tab_button_set_False(self):
        self.ui.clear_button.setEnabled(False)
        self.ui.delete_entry.setEnabled(False)
        self.ui.reset_find_preview.setEnabled(False)
        self.ui.confirm_edit.setEnabled(False)

    def issue_tab_button_set_false(self):
        self.ui.clear_button_2.setEnabled(False)
        self.ui.delete_entry_2.setEnabled(False)
        self.ui.edit_button_2.setEnabled(False)
        self.ui.Reset_issued_preview.setEnabled(False)

    def save_edit(self):
        for i in range(self.ui.find_stock_preview.rowCount()):
            temp_dict = {}
            for j in range(self.ui.find_stock_preview.columnCount()):
                temp_dict[self.ui.find_stock_preview.horizontalHeaderItem(
                    j).text()] = self.ui.find_stock_preview.item(i, j).text()
            index = copy(int(temp_dict["sheet_index"]))
            del(temp_dict["sheet_index"])
            self.stock_sheet[index-1] = deepcopy(temp_dict)
        self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
            self.stock_sheet), self.schema["stock_sheet"]["columns"])

    def delete_entry_from_tabel(self, tab="stock"):
        if tab == "stock":
            if len(self.ui.find_stock_preview.selectedIndexes()) == self.ui.find_stock_preview.columnCount() and self.ui.find_stock_preview.selectedIndexes().count(self.ui.find_stock_preview.selectedIndexes()[0]) == len(self.ui.find_stock_preview.selectedIndexes()):
                # save item
                temp_dict = {}
                for i in range(self.ui.find_stock_preview.columnCount()):
                    temp_dict[self.ui.find_stock_preview.horizontalHeaderItem(i).text()] = self.ui.find_stock_preview.item(
                        self.ui.find_stock_preview.selectedIndexes()[0].row(), i).text()
                self.deleted.append(self.stock_sheet.pop(
                    int(temp_dict["sheet_index"])-1))
                self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
                    self.stock_sheet), self.schema["stock_sheet"]["columns"])
                self.ui.find_stock_preview.removeRow(
                    self.ui.find_stock_preview.selectedIndexes()[0].row())
            else:
                indexes = self.ui.find_stock_preview.selectedIndexes()
                for index in indexes:
                    self.ui.find_stock_preview.setItem(
                        index.row(), index.column(), qtw.QTableWidgetItem(str('')))
                    self.stock_sheet[int(self.ui.find_stock_preview.item(index.row(
                    ), 0).text())-1][self.ui.find_stock_preview.horizontalHeaderItem(index.column()).text()] = ''
                self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
                    self.stock_sheet), self.schema["stock_sheet"]["columns"])

        elif tab == "issue":
            if len(self.ui.issue_stock_preview.selectedIndexes()) == self.ui.issue_stock_preview.columnCount() and self.ui.issue_stock_preview.selectedIndexes().count(self.ui.issue_stock_preview.selectedIndexes()[0]) == len(self.ui.issue_stock_preview.selectedIndexes()):
                # save item
                temp_dict = {}
                for i in range(self.ui.issue_stock_preview.columnCount()):
                    temp_dict[self.ui.issue_stock_preview.horizontalHeaderItem(i).text()] = self.ui.issue_stock_preview.item(
                        self.ui.issue_stock_preview.selectedIndexes()[0].row(), i).text()
                self.deleted.append(self.issue_sheet.pop(
                    int(temp_dict["sheet_index"])-1))
                # self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
                #     self.stock_sheet), self.schema["stock_sheet"]["columns"])
                self.ui.issue_stock_preview.removeRow(
                    self.ui.issue_stock_preview.selectedIndexes()[0].row())
            else:
                indexes = self.ui.issue_stock_preview.selectedIndexes()
                for index in indexes:
                    self.ui.issue_stock_preview.setItem(
                        index.row(), index.column(), qtw.QTableWidgetItem(str('')))
                    self.issue_sheet[int(self.ui.issue_stock_preview.item(index.row(
                    ), 0).text())-1][self.ui.issue_stock_preview.horizontalHeaderItem(index.column()).text()] = ''
                # self.load_data_in_tabel(self.ui.current_stock, self.issue_sheet, len(
                #     self.issue_sheet), self.schema["stock_sheet"]["columns"])

    def find_item_in_stock_sheet(self, sheet):

        if sheet == "stock_sheet":
            self.clear_tabel(self.ui.find_stock_preview)
            self.original = []
            self.ui.current_stock_find_entry.clearSelection()
            # sku_code = self.ui.sku_code.text()
            # if ''.join(sku_code.split()) == "":
            search_dict = {}
            for i in range(len(self.schema["stock_sheet"]["columns"])):
                try:
                    # print(self.ui.current_stock_find_entry.item(0,i).text())
                    if self.ui.current_stock_find_entry.item(0, i).text() != "":
                        search_dict[self.schema["stock_sheet"]["columns"][i]
                                    ] = self.ui.current_stock_find_entry.item(0, i).text()
                except AttributeError:
                    False
                #     print("in exception")
            if not bool(search_dict):
                return
            # search_result = []
            count = 0
            # print(self.ui.current_stock.items())

            for dicti in self.stock_sheet:
                count += 1
                found_flag = True
                for key in list(search_dict):
                    if search_dict[key] != str(dicti[key]):
                        found_flag = False
                        break
                if found_flag == True:
                    dicti_temp = {"sheet_index": count}
                    dicti_temp.update(dicti)
                    self.original.append(dicti_temp)
            if bool(self.original):
                self.ui.find_stock_preview.setEnabled(True)
                self.load_data_in_tabel(self.ui.find_stock_preview, self.original, len(
                    self.original), list(self.original[0]))
                self.ui.clear_button.setEnabled(True)
                self.ui.reset_find_preview.setEnabled(True)
                self.ui.confirm_edit.setEnabled(True)
                # self.ui.edit_button.setEnabled(True)
                self.ui.delete_entry.setEnabled(True)

        elif sheet == "issue_sheet":
            self.clear_tabel(self.ui.issue_stock_preview)
            self.original_issue = []
            self.ui.issued_stock_find_entry.clearSelection()
            # sku_code = self.ui.sku_code.text()
            # if ''.join(sku_code.split()) == "":
            search_dict = {}
            for i in range(len(self.schema["issue_sheet"]["columns"])):
                try:
                    # print(self.ui.current_stock_find_entry.item(0,i).text())
                    if self.ui.issued_stock_find_entry.item(0, i).text() != "":
                        search_dict[self.schema["issue_sheet"]["columns"][i]
                                    ] = self.ui.issued_stock_find_entry.item(0, i).text()
                except AttributeError:
                    False
                #     print("in exception")
            if not bool(search_dict):
                return
            # search_result = []
            count = 0
            # print(self.ui.current_stock.items())

            for dicti in self.issue_sheet:
                count += 1
                found_flag = True
                for key in list(search_dict):
                    if search_dict[key] != str(dicti[key]):
                        found_flag = False
                        break
                if found_flag == True:
                    dicti_temp = {"sheet_index": count}
                    dicti_temp.update(dicti)
                    self.original_issue.append(dicti_temp)
            if bool(self.original_issue):
                self.ui.issue_stock_preview.setEnabled(True)
                self.load_data_in_tabel(self.ui.issue_stock_preview, self.original_issue, len(
                    self.original_issue), list(self.original_issue[0]))
                self.ui.clear_button_2.setEnabled(True)
                self.ui.Reset_issued_preview.setEnabled(True)
                self.ui.edit_button_2.setEnabled(True)
                self.ui.delete_entry_2.setEnabled(True)

    def load_data_in_tabel(self, tabel, values, rowcount, columns, affected_row=None, read_only=[]):
        # affected_row =
        assert(isinstance(tabel, qtw.QTableWidget))
        tabel.setRowCount(rowcount)
        tabel.setColumnCount(
            len(columns))
        tabel.setHorizontalHeaderLabels(
            columns)
        if affected_row == None:
            # load entire data in table
            for row, dicti in enumerate(values):
                for col, key in enumerate(dicti):
                    tabel.setItem(
                        row, col, qtw.QTableWidgetItem(str(dicti[key])))

        else:
            # load only the respective data
            if bool(read_only):
                for col, key in enumerate(values):
                    # here values is one dict
                    item = qtw.QTableWidgetItem(str(values[key]))
                    if key in read_only:
                        item.setFlags(qtc.Qt.ItemIsEditable)
                    tabel.setItem(affected_row, col,
                                  item)
            else:
                for col, key in enumerate(values):
                    # here values is one dict
                    tabel.setItem(affected_row, col,
                                  qtw.QTableWidgetItem(str(values[key])))

        tabel.horizontalHeader(
        ).resizeSections(qtw.QHeaderView.ResizeToContents)

    def save_project(self, save_path):
        self.stock_sheet.to_excel(os.path.join(
            save_path, self.schema["stock_sheet"]["filename"]))
        self.issue_sheet.to_excel(os.path.join(
            save_path, self.schema["issue_sheet"]["filename"]))
        self.issue_return_sheet.to_excel(os.path.join(
            save_path, self.schema["issue_return_sheet"]["filename"]))
        self.sale_sheet.to_excel(os.path.join(
            save_path, self.schema["sale_sheet"]["filename"]))

    def open_project(self, path=False):
        self.ui.actionOpen_Project.setEnabled(False)
        # try:
        if path == False:
            self.project_path = qtw.QFileDialog.getExistingDirectory(
                self, "Select Directory")
        if bool(self.project_path):
            f = open(os.path.join(self.project_path, "schema.json"))
            self.schema = json.loads(f.read())
            f.close()

            self.stock_sheet = pd.read_excel(os.path.join(
                self.project_path, self.schema["stock_sheet"]["filename"])).to_dict(orient="records")
            self.issue_sheet = pd.read_excel(os.path.join(
                self.project_path, self.schema["issue_sheet"]["filename"])).to_dict(orient="records")
            self.issue_return_sheet = pd.read_excel(os.path.join(
                self.project_path, self.schema["issue_return_sheet"]["filename"])).to_dict(orient="records")
            self.sale_sheet = pd.read_excel(os.path.join(
                self.project_path, self.schema["sale_sheet"]["filename"])).to_dict(orient="records")

            # stock tab
            self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
                self.stock_sheet), self.schema["stock_sheet"]["columns"])
            self.load_data_in_tabel(self.ui.current_stock_find_entry, [
            ], 1, self.schema["stock_sheet"]["columns"])
            self.load_data_in_tabel(
                self.ui.find_stock_preview, [], 1, self.schema["stock_sheet"]["columns"])
            self.load_data_in_tabel(self.ui.current_entry_add_stock, [
            ], 1, self.schema["stock_sheet"]["columns"])

            # issue tab
            self.load_data_in_tabel(self.ui.issued_stock_find_entry, [
            ], 1, self.schema["issue_sheet"]["columns"])
            self.load_data_in_tabel(
                self.ui.issue_stock_preview, [], 1, self.schema["issue_sheet"]["columns"])
            self.load_data_in_tabel(self.ui.sku_find, [], 1,
                                    self.schema["stock_sheet"]["columns"])
            self.load_data_in_tabel(self.ui.issue_items, [],
                                    0, self.schema["issue_sheet"]["columns"])

            # qtw.QHeaderView.setSectionResizeMode(qtw.QHeaderView.Stretch)

            self.ui.tabWidget.setTabEnabled(3, True)
            self.ui.tabWidget.setTabEnabled(2, True)
            self.ui.tabWidget.setTabEnabled(1, True)
            self.ui.tabWidget.setTabEnabled(0, True)
            self.ui.actionSave_As.setEnabled(True)
            self.ui.actionSave.setEnabled(True)

            self.ui.issued_stock_find_entry.horizontalHeader(
            ).resizeSections(qtw.QHeaderView.ResizeToContents)
            self.ui.issue_stock_preview.horizontalHeader(
            ).resizeSections(qtw.QHeaderView.ResizeToContents)
            self.ui.sku_find.horizontalHeader().resizeSections(
                qtw.QHeaderView.ResizeToContents)
            self.ui.issue_items.horizontalHeader().resizeSections(
                qtw.QHeaderView.ResizeToContents)

            self.ui.current_stock_find_entry.horizontalHeader(
            ).resizeSections(qtw.QHeaderView.ResizeToContents)
            self.ui.find_stock_preview.horizontalHeader().resizeSections(
                qtw.QHeaderView.ResizeToContents)
            self.ui.current_entry_add_stock.horizontalHeader(
            ).resizeSections(qtw.QHeaderView.ResizeToContents)
            self.ui.current_stock.horizontalHeader().resizeSections(
                qtw.QHeaderView.ResizeToContents)

            # Set Approval No.
            m1 = max(self.issue_sheet, key=lambda x: x["APPROVAL NO."], default={
                     "APPROVAL NO.": 0})
            m2 = max(self.issue_return_sheet, key=lambda x: x["APPROVAL NO."], default={
                     "APPROVAL NO.": 0})
            m1 = max(m1["APPROVAL NO."], m2["APPROVAL NO."])
            if self.schema["approval_no"] < m1:
                self.schema["approval_no"] = m1+1
            self.ui.approval_no_2.setValue(self.schema["approval_no"])
            self.schema["approval_no"] += 1

        # except:
        #     msg = qtw.QMessageBox()
        #     msg.setText("Unable to open project, please check path/schema")
        #     msg.setStandardButtons(qtw.QMessageBox.Ok)
        #     msg.setIcon(qtw.QMessageBox.Warning)
        #     msg.setModal(True)
        #     msg.exec()

    def New_project(self):
        self.project_path = qtw.QFileDialog.getExistingDirectory(
            self, "Select Directory")
        f = open("schema.json")
        self.schema = json.loads(f.read())
        f.close()
        f = open(os.path.join(self.project_path, "schema.json"), "w")
        f.write(json.dumps(self.schema))
        f.close()
        pd.DataFrame(columns=self.schema["stock_sheet"]["columns"]).to_excel(os.path.join(
            self.project_path, self.schema["stock_sheet"]["filename"]), index=False)
        pd.DataFrame(columns=self.schema["issue_sheet"]["columns"]).to_excel(os.path.join(
            self.project_path, self.schema["issue_sheet"]["filename"]), index=False)
        pd.DataFrame(columns=self.schema["issue_return_sheet"]["columns"]).to_excel(os.path.join(
            self.project_path, self.schema["issue_return_sheet"]["filename"]), index=False)
        pd.DataFrame(columns=self.schema["sale_sheet"]["columns"]).to_excel(os.path.join(
            self.project_path, self.schema["sale_sheet"]["filename"]), index=False)

        self.open_project(path=True)
        # self.ui.tabWidget.setTabEnabled(3, True)
        # self.ui.tabWidget.setTabEnabled(2, True)
        # self.ui.tabWidget.setTabEnabled(1, True)
        # self.ui.tabWidget.setTabEnabled(0, True)
        # self.ui.actionSave_As.setEnabled(True)
        # self.ui.actionSave.setEnabled(True)

    def exit(self):
        if self.saved == False:
            msg = qtw.QMessageBox()
            msg.setText("Please save before closing")
            msg.setStandardButtons(qtw.QMessageBox.Save |
                                   qtw.QMessageBox.Discard)
            msg.setModal(True)
            msg.buttonClicked.connect(self.save_warning)
            msg.exec_()

            # print(retval)
        else:
            os._exit(0)

    def save_warning(self, item):
        if item.text() == "Save":
            self.save_project(self.project_path)


if __name__ == "__main__":

    app = qtw.QApplication([])

    window = MainWindow()
    window.showMaximized()

    app.exec()
