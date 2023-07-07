# This Python file uses the following encoding: utf-8
import sys,xlrd

from PySide6.QtWidgets import QApplication, QWidget
from PyQt6.QtWidgets import QApplication, QDialog, QMainWindow, QMessageBox, QPushButton , QVBoxLayout, QToolBar
from PyQt6.QtGui import QKeySequence
from fpdf import FPDF
import mysql.connector as mc
import pandas as pd
# Important:
# You need to run the following command to generate the ui_form.py file
#     pyside6-uic form.ui -o ui_form.py, or
#     pyside2-uic form.ui -o ui_form.py
from ui_form import Ui_Widget


class Widget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_Widget()
        self.ui.setupUi(self)
        self.ui.add_item.clicked.connect(self.tambah_data)
        self.ui.update.clicked.connect(self.update_data)
        self.ui.delete_item.clicked.connect(self.hapus_data)
        self.ui.clear_all.clicked.connect(self.clear_data)
        self.grab_all()
        self.ui.ekspor.clicked.connect(self.ekspor_data)
        self.ui.ekspor_pdf.clicked.connect(self.export_pdf)
        self.ui.impor.clicked.connect(self.import_data)
        self.ui.exit.clicked.connect(self.keluar)

    # tambah data pada list
    def tambah_data(self):
        item = self.ui.additem_line_edit.text()
        if item != "" :
            self.ui.my_list.addItem(item)
            self.ui.additem_line_edit.setText("")

        try:
            mydb = mc.connect(
                host="localhost",
                user="root",
                password="",
                database="to_do"
            )
            mycursor = mydb.cursor()
            sql = "INSERT INTO list (item) VALUES (%s)"
            val = (item,)

            mycursor.execute(sql, val)
            mydb.commit()

            mydb.close()
            self.ui.keterangan.setText("Data berhasil ditambahkan")

        except:
            print("Koneksi Gagal")

    #Update data yang sedang diseleksi
    def update_data(self):
        item = self.ui.additem_line_edit.text()
        select = self.ui.my_list.currentItem().text()
        if item != "" :
            try:
                mydb = mc.connect(
                    host="localhost",
                    user="root",
                    password="",
                    database="to_do"
                )
                mycursor = mydb.cursor()
                sql = "UPDATE list SET item = %s WHERE item = %s"
                val = (item,select,)
                mycursor.execute(sql, val)
                mydb.commit()
                mydb.close()
                self.ui.my_list.currentItem().setText(item)
                self.ui.additem_line_edit.setText("")
                self.ui.keterangan.setText("Data berhasil diupdate")

            except:
                print("Koneksi Gagal")

    # hapus satu data yang ada di list
    def hapus_data(self):
        itemz = self.ui.my_list.currentItem().text()

        try:
            mydb = mc.connect(
                host="localhost",
                user="root",
                password="",
                database="to_do"
            )
            mycursor = mydb.cursor()
            sql = """DELETE FROM list where item = %s"""
            val = (itemz,)

            mycursor.execute(sql, val)
            mydb.commit()
            mydb.close()
            self.ui.keterangan.setText("Data berhasil dihapus")
            clicked = self.ui.my_list.currentRow()
            self.ui.my_list.takeItem(clicked)

        except:
            print("Koneksi Gagal")

    # hapus semua data di list
    def clear_data(self):
        self.ui.my_list.clear()

    # tambah data ke database
    def save_data(self):
        try:
            mydb = mc.connect(
                host="localhost",
                user="root",
                password="",
                database="to_do"
            )
            mycursor = mydb.cursor()
            mycursor.execute("DELETE FROM list")

        except:
            print("Koneksi Gagal")

        items = []
        for index in range(self.ui.my_list.count()):
            items.append(self.ui.my_list.item(index))

        sql = "INSERT INTO list (item) VALUES (%s)"

        for item in items:
#           print(item.text())
            val = (item.text(),)
            # Execute sql Query
            mycursor.execute(sql, val)

        mydb.commit()
        mydb.close()
        self.ui.keterangan.setText("Data berhasil disimpan ke database")

    # menampilkan data dari database ke dalam list
    def grab_all(self):
        try:
            mydb = mc.connect(
                host="localhost",
                user="root",
                password="",
                database="to_do"
            )
            mycursor = mydb.cursor()
            mycursor.execute("SELECT item FROM list")
            records = mycursor.fetchall()

            mydb.commit()
            mydb.close()

            for record in records:
                self.ui.my_list.addItem(str(record[0]))

        except:
            print("Koneksi Gagal")

    # ekspor data dari database to_do tabel list kedalam file excel
    def ekspor_data(self):
        # Konek ke Basis Data
        mydb = mc.connect(
            host="localhost",
            user="root",
            password="",
            database="to_do"
        )
        mycursor = mydb.cursor()
        mycursor.execute("SELECT item FROM list")

        rows = mycursor.fetchall()
        df = pd.DataFrame(rows)

        col_names= ["Kegiatan"]
        df.columns = col_names
        df.to_excel('to_do.xlsx', index=False)

        mycursor.close()
        mydb.close()
        self.ui.keterangan.setText("Data berhasil diekspor dalam bentuk excel")

    # import data dari file excel ke database
    def import_data(self):
        book = xlrd.open_workbook("to_do.xlsx")
        sheet = book.sheet_by_name("Sheet1")
        mydb = mc.connect(
            host="localhost",
            user="root",
            password="",
            database="to_do"
        )
        mycursor = mydb.cursor()
        sql = "INSERT INTO list (item) VALUES (%s)"
        for r in range(1, sheet.nrows):
            keg            = sheet.cell(r,0).value
            val = (keg,)
            # Execute sql Query
            mycursor.execute(sql, val,)
            mydb.commit()

        mycursor.close()
        mydb.commit()
        mydb.close()
        self.ui.keterangan.setText("Data berhasil diimpor dari file excel")
        self.grab_all()

    # ekspor data dari database to_do tabel list kedalam file pdf
    def export_pdf(self):
        my_pdf = FPDF()
        my_pdf.add_page()
        my_pdf.set_font("Arial", size=17)

        items = []
        for index in range(self.ui.my_list.count()):
            items.append(self.ui.my_list.item(index))

        my_pdf.cell(200, 10, txt="Aplikasi To Do List Reza", ln=1, align="C")
        my_pdf.set_font("Arial", size=15)
        my_pdf.cell(200, 10, txt="List Kegiatan : ", ln=2, align="L")
        my_pdf.set_font("Arial", size=13)
        nom=0
        for item in items:
            nom+=1
            my_pdf.cell(200, 10, txt=str(nom)+ ". "+item.text(), ln=1, align="L")

        my_pdf.cell(200, 10, txt="", ln=2, align="L")
        my_pdf.cell(200, 10, txt="", ln=2, align="L")
        my_pdf.set_font("Arial", size=10)
        my_pdf.cell(200, 10, txt="'Produktif Adalah Mendahulukan Apa Yang Harus Kita Kerjakan'", ln=2, align="C")

        my_pdf.output("To Do.pdf")
        self.ui.keterangan.setText("Data berhasil diekspor dalam bentuk pdf")

    def keluar(self):
        self.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    widget = Widget()
    widget.show()
    sys.exit(app.exec())
