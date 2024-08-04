import sqlite3
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QTextEdit, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox, QComboBox, QGridLayout, QSizePolicy
from PySide6.QtGui import QColor

# 连接到SQLite数据库（如果数据库不存在，则会创建一个）
conn = sqlite3.connect('inventory.db')

# 创建一个游标对象
cursor = conn.cursor()

# 创建items表
cursor.execute('''
CREATE TABLE IF NOT EXISTS items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code INTEGER NOT NULL,
    name TEXT,
    model TEXT,
    package TEXT,
    purchase_price REAL,
    quantity INTEGER,
    remaining_quantity INTEGER,
    storage_location TEXT,
    purchase_link TEXT,
    project TEXT,
    notes TEXT
)
''')

# 提交更改并关闭连接
conn.commit()
conn.close()

class InventoryApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("库存管理软件-PlayerPencil制作")
        self.setGeometry(300, 300, 1200, 800)
        
        self.layout = QVBoxLayout()
        
        # 创建一个容器来包裹表单区域和其他元素
        self.form_container = QWidget()
        self.form_layout = QVBoxLayout(self.form_container)
        
        # 输入区域
        self.input_layout = QGridLayout()
        self.fields = [
            ("器件名称", "name"),
            ("器件型号", "model"),
            ("封装", "package"),
            ("购买价格", "purchase_price"),
            ("数量", "quantity"),
            ("剩余数量", "remaining_quantity"),
            ("存储位置", "storage_location"),
            ("购买链接", "purchase_link"),
            ("需求项目", "project"),
            ("备注", "notes")
        ]
        self.inputs = {}
        self.name_input = 0
        for idx, (label_text, field_name) in enumerate(self.fields):
            label = QLabel(label_text)
            if field_name == "name":
                input_field = QComboBox()
                self.name_input = input_field
                self.populate_name_combobox(input_field)
                input_field.setEditable(True)
                self.input_layout.addWidget(label, idx // 3 * 2, idx % 3)
                self.input_layout.addWidget(input_field, idx // 3 * 2 + 1, idx % 3)
            elif field_name == "package":
                input_field = QComboBox()
                input_field.setEditable(True)
                self.input_layout.addWidget(label, idx // 3 * 2, idx % 3)
                self.input_layout.addWidget(input_field, idx // 3 * 2 + 1, idx % 3)
                self.inputs[field_name] = input_field
                self.inputs["name"].currentTextChanged.connect(lambda: self.populate_package_combobox(self.inputs["package"], self.inputs["name"].currentText()))
            elif field_name == "notes":
                input_field = QTextEdit()
                input_field.setFixedHeight(80)  # 设置为四行高
                self.input_layout.addWidget(label, idx // 3 * 2, 0)
                self.input_layout.addWidget(input_field, idx // 3 * 2 + 1, 0, 1, 3)
            elif field_name == "storage_location":
                input_field = QComboBox()
                input_field.setEditable(True)
                self.input_layout.addWidget(label, idx // 3 * 2, idx % 3)
                self.input_layout.addWidget(input_field, idx // 3 * 2 + 1, idx % 3)
                self.inputs[field_name] = input_field
                self.inputs["name"].currentTextChanged.connect(lambda: self.populate_storage_location_combobox(self.inputs["storage_location"], self.inputs["name"].currentText()))
            elif field_name == "project":
                input_field = QComboBox()
                input_field.setEditable(True)
                self.populate_project_combobox(input_field)
                self.input_layout.addWidget(label, idx // 3 * 2, idx % 3)
                self.input_layout.addWidget(input_field, idx // 3 * 2 + 1, idx % 3)
            else:
                input_field = QLineEdit()
                self.input_layout.addWidget(label, idx // 3 * 2, idx % 3)
                self.input_layout.addWidget(input_field, idx // 3 * 2 + 1, idx % 3)
            self.inputs[field_name] = input_field
        
        self.add_button = QPushButton("添加")
        self.add_button.setStyleSheet("background-color: lightgreen;")
        self.add_button.clicked.connect(self.add_item)
        self.clear_button = QPushButton("清除")
        self.clear_button.setStyleSheet("background-color: lightblue;")
        self.clear_button.clicked.connect(self.clear_inputs)
        self.delete_button = QPushButton("删除")
        self.delete_button.setStyleSheet("background-color: lightcoral;")
        self.delete_button.clicked.connect(self.delete_item)
        
        # 将按钮单独放在一行
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.clear_button)
        button_layout.addWidget(self.delete_button)
        
        self.input_layout.addLayout(button_layout, (len(self.fields) // 3 + 1) * 2, 0, 1, 3)
        
        self.form_layout.addLayout(self.input_layout)
        
        # 查找区域
        self.search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_button = QPushButton("查找")
        self.search_button.clicked.connect(self.search_item)
        self.search_layout.addWidget(QLabel("查找器件:"))
        self.search_layout.addWidget(self.search_input)
        self.search_layout.addWidget(self.search_button)
        
        self.form_layout.addLayout(self.search_layout)
        
        # 设置容器的最大高度
        self.form_container.setMaximumHeight(340)
        
        self.layout.addWidget(self.form_container)
        
        # 表格区域
        self.table = QTableWidget()
        self.table.setColumnCount(len(self.fields) + 2)
        self.table.setHorizontalHeaderLabels(["ID", "器件编码"] + [label for label, _ in self.fields])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.cellClicked.connect(self.load_item)
        self.table.itemChanged.connect(self.update_table_item)
        
        # 设置表格的大小策略为可伸缩
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        self.layout.addWidget(self.table)
        
        self.setLayout(self.layout)
        
        self.load_data()
    
    def populate_name_combobox(self, combobox):
        conn = sqlite3.connect('inventory.db')
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT name FROM items")
        names = cursor.fetchall()
        for name in names:
            combobox.addItem(name[0])
        conn.close()
    
    def populate_storage_location_combobox(self, combobox, name):
        combobox.clear()
        if name:
            conn = sqlite3.connect('inventory.db')
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT storage_location FROM items WHERE name = ?", (name,))
            locations = cursor.fetchall()
            for location in locations:
                combobox.addItem(location[0])
            conn.close()
    
    def populate_project_combobox(self, combobox):
        combobox.clear()
        conn = sqlite3.connect('inventory.db')
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT project FROM items")
        projects = cursor.fetchall()
        for project in projects:
            combobox.addItem(project[0])
        conn.close()
    
    # 添加 populate_package_combobox 方法
    def populate_package_combobox(self, combobox, name):
        combobox.clear()
        if name:
            conn = sqlite3.connect('inventory.db')
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT package FROM items WHERE name = ?", (name,))
            packages = cursor.fetchall()
            for package in packages:
                combobox.addItem(package[0])
            conn.close()
    
    def load_data(self):
        conn = sqlite3.connect('inventory.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM items")
        rows = cursor.fetchall()
        
        self.table.setRowCount(len(rows))
        for row_idx, row_data in enumerate(rows):
            for col_idx, col_data in enumerate(row_data):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(col_data)))
        
        conn.close()

    def clear_inputs(self):
        for field in self.inputs.values():
            if isinstance(field, QLineEdit) or isinstance(field, QComboBox):
                field.clear()
            elif isinstance(field, QTextEdit):
                field.clear()
        self.populate_name_combobox(self.name_input)
        

    def update_table_item(self, item):
        row = item.row()
        col = item.column()
        new_value = item.text()
        item_id = self.table.item(row, 0).text()

        conn = sqlite3.connect('inventory.db')
        cursor = conn.cursor()

        # 根据列索引更新相应的字段
        if col == 2:
            cursor.execute("UPDATE items SET name = ? WHERE id = ?", (new_value, item_id))
        elif col == 3:
            cursor.execute("UPDATE items SET model = ? WHERE id = ?", (new_value, item_id))
        elif col == 4:
            cursor.execute("UPDATE items SET package = ? WHERE id = ?", (new_value, item_id))
        elif col == 5:
            cursor.execute("UPDATE items SET purchase_price = ? WHERE id = ?", (new_value, item_id))
        elif col == 6:
            cursor.execute("UPDATE items SET quantity = ? WHERE id = ?", (new_value, item_id))
        elif col == 7:
            cursor.execute("UPDATE items SET remaining_quantity = ? WHERE id = ?", (new_value, item_id))
        elif col == 8:
            cursor.execute("UPDATE items SET storage_location = ? WHERE id = ?", (new_value, item_id))
        elif col == 9:
            cursor.execute("UPDATE items SET purchase_link = ? WHERE id = ?", (new_value, item_id))
        elif col == 10:
            cursor.execute("UPDATE items SET project = ? WHERE id = ?", (new_value, item_id))
        elif col == 11:
            cursor.execute("UPDATE items SET notes = ? WHERE id = ?", (new_value, item_id))

        conn.commit()
        conn.close()
    
    def add_item(self):
        item_data = {
        field: self.inputs[field].currentText() if isinstance(self.inputs[field], QComboBox) else self.inputs[field].text() if field != "notes" else self.inputs[field].toPlainText()
        for _, field in self.fields
        }
        
        conn = sqlite3.connect('inventory.db')
        cursor = conn.cursor()
        cursor.execute("SELECT MAX(code) FROM items")
        max_code = cursor.fetchone()[0]
        item_data['code'] = (max_code + 1) if max_code is not None else 1
        
        cursor.execute('''
            INSERT INTO items (code, name, model, package, purchase_price, quantity, remaining_quantity, storage_location, purchase_link, project, notes)
            VALUES (:code, :name, :model, :package, :purchase_price, :quantity, :remaining_quantity, :storage_location, :purchase_link, :project, :notes)
        ''', item_data)
        conn.commit()
        conn.close()
        
        self.load_data()
        for input_field in self.inputs.values():
            if isinstance(input_field, QComboBox):
                input_field.setCurrentIndex(-1)
            else:
                input_field.clear()
    
    # load_item 方法
    def load_item(self, row, column):
        for col_idx, (_, field) in enumerate(self.fields):
            if field == "notes":
                self.inputs[field].setPlainText(self.table.item(row, col_idx + 2).text())
            elif isinstance(self.inputs[field], QComboBox):
                self.inputs[field].setCurrentText(self.table.item(row, col_idx + 2).text())
            else:
                self.inputs[field].setText(self.table.item(row, col_idx + 2).text())
        self.current_item_id = self.table.item(row, 0).text()

    # update_item 方法
    def update_item(self):
        name = self.inputs["name"].currentText()
        package = self.inputs["package"].currentText()
        notes = self.inputs["notes"].toPlainText()
        storage_location = self.inputs["storage_location"].currentText()
        project = self.inputs["project"].currentText()

        conn = sqlite3.connect('inventory.db')
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE items
            SET package = ?, notes = ?, storage_location = ?, project = ?
            WHERE name = ?
        """, (package, notes, storage_location, project, name))
        conn.commit()
        conn.close()

        # 刷新界面上的数据
        self.populate_package_combobox(self.inputs["package"], name)
        self.populate_storage_location_combobox(self.inputs["storage_location"], name)
        self.populate_project_combobox(self.inputs["project"])
    
    def delete_item(self):
        if hasattr(self, 'current_item_id'):
            conn = sqlite3.connect('inventory.db')
            cursor = conn.cursor()
            cursor.execute('DELETE FROM items WHERE id = ?', (self.current_item_id,))
            conn.commit()
            conn.close()
            
            self.load_data()
            for input_field in self.inputs.values():
                if isinstance(input_field, QComboBox):
                    input_field.setCurrentIndex(-1)
                else:
                    input_field.clear()
            del self.current_item_id
    
    def search_item(self):
        keyword = self.search_input.text()
        conn = sqlite3.connect('inventory.db')
        cursor = conn.cursor()
        
        if keyword:
            query = f"SELECT * FROM items WHERE name LIKE ? OR model LIKE ? OR package LIKE ? OR storage_location LIKE ? OR project LIKE ? OR notes LIKE ?"
            cursor.execute(query, (f'%{keyword}%',) * 6)
            rows = cursor.fetchall()
            
            if not rows:
                QMessageBox.information(self, "查找结果", "查找不到该器件！")
                cursor.execute("SELECT * FROM items")
                rows = cursor.fetchall()
            
            # 排序匹配的项到前面
            rows.sort(key=lambda row: any(keyword.lower() in str(cell).lower() for cell in row), reverse=True)
        else:
            # QMessageBox.warning(self, "输入错误", "请输入查找器件。")
            cursor.execute("SELECT * FROM items")
            rows = cursor.fetchall()
        
        self.table.setRowCount(len(rows))
        for row_idx, row_data in enumerate(rows):
            for col_idx, col_data in enumerate(row_data):
                item = QTableWidgetItem(str(col_data))
                if keyword and keyword.lower() in str(col_data).lower():
                    item.setBackground(QColor('yellow'))
                self.table.setItem(row_idx, col_idx, item)
        
        conn.close()

if __name__ == "__main__":
    app = QApplication([])
    window = InventoryApp()
    window.show()
    app.exec()