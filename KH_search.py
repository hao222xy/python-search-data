import sys
from PyQt5.QtCore import QDate, QDateTime
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import time
from sqlalchemy import create_engine
import pandas as pd
from PyQt5 import QtWidgets, QtCore

class Ui_MainWindow (QWidget):
    def __init__(self):
        super(Ui_MainWindow, self).__init__()
        self.initUI()

    def initUI(self):
        # 设置标题与初始大小
        self.setWindowTitle('坤厚订单信息查询')
        self.resize(1200, 750)


        # 垂直布局/水平布局 QVBoxLayout/QHBoxLayout
        layout = QVBoxLayout()
        self.setLayout(layout)

        # 创建第一个日期时间空间,并把当前日期时间赋值,并修改显示格式
        self.label1 = QLabel('下发时间')
        self.dateEdit1 = QDateTimeEdit(QDateTime.currentDateTime(), self)
        self.dateEdit1.setDisplayFormat('yyyy-MM-dd')

        # 构造一个QTableWidget对象，设置表格为9行14列
        TableWidget = QTableWidget(8, 14)

        # 设置第一个日期最大值与最小值，在当前日期的基础上，后一年与前一年
        self.dateEdit1.setMinimumDate(QDate.currentDate().addDays(-365))
        self.dateEdit1.setMaximumDate(QDate.currentDate().addDays(365))

        # 设置水平方向的表头标签与垂直方向上的表头标签，注意必须在初始化行列之后进行，否则，没有效果
        TableWidget.setHorizontalHeaderLabels(['编号', '订单编号', 'AGV型号', '启动点', '启动申请点', '任务点', '任务申请点', '优先级', '任务状态', '任务模式', '任务类型', '系统模式', '下发时间',
             '完成时间'])

        # 设置表格头为伸缩模式
        TableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # 表格整行选中
        TableWidget.setSelectionBehavior(QAbstractItemView.SelectRows)

        # 设置控件允许弹出
        self.dateEdit1.setCalendarPopup(True)

        # 创建按钮并绑定一个自定义槽函数
        self.btn1 = QPushButton('点击查询')
        self.btn1.clicked.connect(self.onButtonClick)

        self.btn2 = QPushButton('导出xlsx')
        self.btn2.clicked.connect(self.onButtonClick)


        # 布局控件的加载与设置,可加载多个控件
        layout.addWidget(self.label1)
        layout.addWidget(self.dateEdit1)

        layout.addWidget(self.btn1)

        layout.addWidget(TableWidget)
        self.setLayout(layout)

        layout.addWidget(self.btn2)
        self.setLayout(layout)

    def onButtonClick(self):
        # dateTime是QDateTimeEdit的一个方法，返回QDateTime时间格式
        # 需要再用toPyDateTime转变回python的时间格式
        dateTime1 = str(self.dateEdit1.dateTime().toPyDateTime())[0:10]

        ##### MySQL数据库 —> xlsx文件
        # 1.创建数据库连接
        engine = create_engine('mysql+pymysql://root:123456@localhost/wcs_emerson')
        # 2.读取MySQL数据
        df1 = pd.read_sql(
            sql="SELECT * FROM wcs_emerson.kh_report_order_info where ENTERDATE LIKE '" + dateTime1 + "%%'", con=engine)
        # 3.导出csv文件
        df1.to_csv('data.csv', header=None, index=False)
        df2 = pd.read_csv('data.csv',
                          names=['编号', '订单编号', 'AGV型号', '启动点', '启动申请点', '任务点', '任务申请点', '优先级', '任务状态', '任务模式', '任务类型',
                                 '系统模式', '下发时间','完成时间'])
        df2.to_csv('data.csv', index=False)

        # 4.使用pandas将csv文件转成xlsx文件
        def csv_to_xlsx_pd():
            csv = pd.read_csv('data.csv', encoding='utf-8')
            csv.to_excel('data.xlsx', sheet_name='data')

        if __name__ == '__main__':
            csv_to_xlsx_pd()

        # python时间格式转换
        n_time11 = time.strptime(dateTime1, "%Y-%m-%d")
        n_time1 = int(time.strftime('%Y%m%d', n_time11))

    # 显示已有数据，并且添加网可见  查询
        showDatas = str()
        engine = create_engine()
        cur = engine.cursor()
        sql="SELECT * FROM wcs_emerson.kh_report_order_info'" + showDatas + "%%'"
        cur.execute(sql)
        data = cur.fetchall()

        if data:
            """
            要获取当前表格部件中的行数，可以通过rowCount()方法获取，
            要设置表格部件的行数，可以通过setRowCount（int rows）调整表格的行数，
            如果参数rows小于现在表格中的实际行数，则表格中超出参数的行数数据会丢弃，
            就算是后面将行数或列数恢复也不能恢复相关数据
            """
            showDatas.setRowCount(0)#获取行数
            showDatas.insertRow(0)#插入o行
            for row,form in enumerate(data):
                for column,item in enumerate(form):
                    showDatas.setItem(row,column,QTableWidgetItem(str(item)))#行，列，赋值
                    column += 1
                row_postition = showDatas.rowCount()
                showDatas.insertRow(row_postition)


# if __name__ == '__main__'的作用是为了防止其他脚本只是调用该类时才开始加载，优化内存使用
if __name__ == '__main__':
    # 调用
    app = QApplication(sys.argv)
    MainWindow = Ui_MainWindow()

    MainWindow.show()
    sys.exit(app.exec_())


