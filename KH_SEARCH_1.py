import sys
from PyQt5.QtCore import QDate, QDateTime
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import time
from sqlalchemy import create_engine
import pandas as pd
from PyQt5 import QtWidgets, QtCore
import pymysql


class Ui_MainWindow (QWidget):
    def __init__(self):
        super(Ui_MainWindow, self).__init__()
        self.initUI()

    def initUI(self):
        # 设置标题与初始大小
        self.setWindowTitle('坤厚订单信息查询')
        self.resize(1200, 750)

        # 构建QTableWidget对象，设置表格行列
        self.order_table = QTableWidget()

        # 垂直布局/水平布局 QVBoxLayout/QHBoxLayout
        layout = QVBoxLayout()
        self.setLayout(layout)

        # 创建第一个日期时间空间,并把当前日期时间赋值,并修改显示格式
        self.label1 = QLabel('下发时间')
        self.dateEdit1 = QDateTimeEdit(QDateTime.currentDateTime(), self)
        self.dateEdit1.setDisplayFormat('yyyy-MM-dd')

        # 设置第一个日期最大值与最小值，在当前日期的基础上，后一年与前一年
        self.dateEdit1.setMinimumDate(QDate.currentDate().addDays(-365))
        self.dateEdit1.setMaximumDate(QDate.currentDate().addDays(365))

        # 创建按钮并绑定一个自定义槽函数
        self.btn1 = QPushButton('点击查询')
        #self.btn1.clicked.connect(self.onButtonClick)
        self.btn1.clicked.connect(self.showOut)

        self.btn2 = QPushButton('导出xlsx')

        # 布局控件的加载与设置,可加载多个控件
        layout.addWidget(self.label1)
        layout.addWidget(self.dateEdit1)
        layout.addWidget(self.btn1)
        layout.addWidget(self.order_table)
        layout.addWidget(self.btn2)
        self.setLayout(layout)

    def onButtonClick(self):
        # dateTime是QDateTimeEdit的一个方法，返回QDateTime时间格式
        # 需要再用toPyDateTime转变回python的时间格式
        dateTime1 = str(self.dateEdit1.dateTime().toPyDateTime())[0:10]
        print(1)

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

        # 设置控件允许弹出
        self.dateEdit1.setCalendarPopup(True)

    def showOut(self):
        dateTime2 = str(self.dateEdit1.dateTime().toPyDateTime())[0:10]
        print(dateTime2)
        conn = pymysql.connect(host='localhost',port=3306,user='root',passwd='123456',db='wcs_emerson',charset='utf8')
        cur = conn.cursor()
        sql = 'select * from kh_report_order_info where ENTERDATE LIKE \'' + dateTime2 + '%\''
        print(sql)
        cur.execute(sql)
        rows = cur.fetchall()
        row = cur.rowcount  # 取得记录个数，用于设置表格的行数
        vol = len(rows[0])  # 取得字段数，用于设置表格的列数
        cur.close()     # 关闭游标
        conn.close()    # 关闭数据库连接

        # 设置表格行列
        self.order_table.setRowCount(row)
        self.order_table.setColumnCount(vol)

        # 设置水平方向的表头标签与垂直方向上的表头标签，注意必须在初始化行列之后进行，否则，没有效果
        self.order_table.setHorizontalHeaderLabels(
            ['编号', '订单编号', 'AGV型号', '启动点', '启动申请点', '任务点', '任务申请点', '优先级', '任务状态', '任务模式', '任务类型', '系统模式', '下发时间',
             '完成时间'])

        # 设置表格头为伸缩模式
        self.order_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # 表格整行选中
        self.order_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        # 表格表头的显示与隐藏
        self.order_table.verticalHeader().setVisible(False)
        #self.order_table.horizontalHeader().setVisible(False)

        for i in range(row):
            for j in range(vol):
                if(j == 0):
                    temp_data = ""  # 临时记录，不能直接插入表格
                    data = QTableWidgetItem(str(temp_data))  # 转换后可插入表格
                else:
                    temp_data = rows[i][j-1]  # 临时记录，不能直接插入表格
                    data = QTableWidgetItem(str(temp_data))  # 转换后可插入表格
                self.order_table.setItem(i, j, data)



# if __name__ == '__main__'的作用是为了防止其他脚本只是调用该类时才开始加载，优化内存使用
if __name__ == '__main__':
    # 调用
    app = QApplication(sys.argv)
    MainWindow = Ui_MainWindow()

    MainWindow.show()
    sys.exit(app.exec_())


