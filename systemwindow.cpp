// systemwindow.cpp源文件
// 功能说明：构建系统窗口，实现系统中不同组件的具体功能


#include "ui_systemwindow.h"
#include <QMessageBox>
#include <QStringListModel>
#include <QCloseEvent>
#include <QFileDialog>
#include <QAxObject>
#include <QProgressDialog>
#include "systemwindow.h"
#include "fileFunction.h"
#include "dataFunction.h"

SystemWindow::SystemWindow(QWidget *parent)
    : QMainWindow(parent) // 窗口
    , ui(new Ui::SystemWindow) // ui界面指针
    , manager(nullptr) // 国旗班制表管理器指针
    , flagGroup() // 国旗班成员容器变量
    , currentSelectedPerson(nullptr) // 保存当前用户选中的队员标签指针
    , isShowingInfo(false) // 标志位，用于区分展示信息和用户主动修改
{
    ui->setupUi(this);
    // 将“使用说明”界面的 QTextEdit 文本框设置为只读模式
    ui->instructionText->setReadOnly(true);

    // 在窗口启动时读取文件
    FlagGroupFileManager::loadFromFile(flagGroup, filename);

    // 执勤管理界面
    // 连接按钮和复选框的信号与槽
    connect(ui->tabulateButton, &QPushButton::clicked, this, &SystemWindow::onTabulateButtonClicked); // 排表按钮点击事件
    connect(ui->clearButton, &QPushButton::clicked, this, &SystemWindow::onClearButtonClicked); // 清空表格按钮点击事件
    connect(ui->alterButton, &QPushButton::clicked, this, &SystemWindow::onResetButtonClicked); // 重置队员执勤总次数按钮点击事件
    connect(ui->deriveButton, &QPushButton::clicked, this, &SystemWindow::onExportButtonClicked); // 导出表格按钮点击事件
    connect(ui->times_rule, &QCheckBox::clicked, this, &SystemWindow::onTotalTimesRuleClicked); // 总次数规则按钮点击事件
    // 连接单选按钮信号与槽
    // 交接工作单选按钮点击事件
    connect(ui->No_handover_rule_radioButton, &QRadioButton::clicked, this, &SystemWindow::onRadioButtonClicked); // 不采用交接规则
    connect(ui->Monday_handover_rule_radioButton, &QRadioButton::clicked, this, &SystemWindow::onRadioButtonClicked); // 仅周二的南鉴湖升旗采用交接规则
    connect(ui->All_handover_rule_radioButton, &QRadioButton::clicked, this, &SystemWindow::onRadioButtonClicked); // 全周（周二至周五）南鉴湖升旗采用交接规则
    // 将不采用交接按钮默认设置为选定状态
    ui->No_handover_rule_radioButton->setChecked(true);

    // 队员管理界面
    // 更新四个组的队员标签界面
    for (int groupIndex = 1; groupIndex <= 4; ++groupIndex) {
        updateListView(groupIndex);
    }
    // 将 FlagGroup 中所有队员的 iswork 信息全部调成 false，对应全组执勤按钮的未选定状态
    for (int groupIndex = 1; groupIndex <= 4; ++groupIndex) {
        auto& members = flagGroup.getGroupMembers(groupIndex); // 返回对应组的队员列表
        for (auto& member : members) {
            member.setIsWork(false); // 修改iswork信息
        }
    }
    // 设置 availableTime_groupBox 中除“全选”按钮外的按钮为 checkable
    QList<QAbstractButton*> buttons = ui->availableTime_groupBox->findChildren<QAbstractButton*>();
    for (QAbstractButton* button : buttons) {
        if (!button->text().contains("全选")) {
            button->setCheckable(true);
        }
    }
    // 连接“全选”按钮的点击事件
    QList<QAbstractButton*> allSelectButtons = ui->availableTime_groupBox->findChildren<QAbstractButton*>();
    for (QAbstractButton* button : allSelectButtons) {
        if (button->text() == "全选") {
            connect(button, &QAbstractButton::clicked, this, &SystemWindow::onAllSelectButtonClicked);
        }
    }
    // 连接一组的信号与槽
    connect(ui->group1_add_pushButton, &QPushButton::clicked, [this]() { onGroupAddButtonClicked(1); });
    connect(ui->group1_delete_pushButton, &QPushButton::clicked, [this]() { onGroupDeleteButtonClicked(1); });
    connect(ui->group1_iswork_radioButton, &QRadioButton::clicked, [this]() { onGroupIsWorkRadioButtonClicked(1); });
    connect(ui->group1_info_listView, &QListView::clicked, [this](const QModelIndex &index) { onListViewItemClicked(index, 1); });
    // 连接二组的信号与槽
    connect(ui->group2_add_pushButton, &QPushButton::clicked, [this]() { onGroupAddButtonClicked(2); });
    connect(ui->group2_delete_pushButton, &QPushButton::clicked, [this]() { onGroupDeleteButtonClicked(2); });
    connect(ui->group2_iswork_radioButton, &QRadioButton::clicked, [this]() { onGroupIsWorkRadioButtonClicked(2); });
    connect(ui->group2_info_listView, &QListView::clicked, [this](const QModelIndex &index) { onListViewItemClicked(index, 2); });
    // 连接三组的信号与槽
    connect(ui->group3_add_pushButton, &QPushButton::clicked, [this]() { onGroupAddButtonClicked(3); });
    connect(ui->group3_delete_pushButton, &QPushButton::clicked, [this]() { onGroupDeleteButtonClicked(3); });
    connect(ui->group3_iswork_radioButton, &QRadioButton::clicked, [this]() { onGroupIsWorkRadioButtonClicked(3); });
    connect(ui->group3_info_listView, &QListView::clicked, [this](const QModelIndex &index) { onListViewItemClicked(index, 3); });
    // 连接四组的信号与槽
    connect(ui->group4_add_pushButton, &QPushButton::clicked, [this]() { onGroupAddButtonClicked(4); });
    connect(ui->group4_delete_pushButton, &QPushButton::clicked, [this]() { onGroupDeleteButtonClicked(4); });
    connect(ui->group4_iswork_radioButton, &QRadioButton::clicked, [this]() { onGroupIsWorkRadioButtonClicked(4); });
    connect(ui->group4_info_listView, &QListView::clicked, [this](const QModelIndex &index) { onListViewItemClicked(index, 4); });

    // 基础信息栏内容
    // 连接 group_combobox 的 currentIndexChanged 信号，组别修改事件
    connect(ui->group_combobox, QOverload<int>::of(&QComboBox::currentIndexChanged), this, &SystemWindow::onGroupComboBoxChanged);
    // 连接队员信息输入框的 editingFinished 信号到编辑结束槽函数，但不包括 group_combobox
    connect(ui->name_lineEdit, &QLineEdit::editingFinished, this, &SystemWindow::onInfoLineEditChanged);
    connect(ui->phone_lineEdit, &QLineEdit::editingFinished, this, &SystemWindow::onInfoLineEditChanged);
    connect(ui->nativePlace_lineEdit, &QLineEdit::editingFinished, this, &SystemWindow::onInfoLineEditChanged);
    connect(ui->school_lineEdit, &QLineEdit::editingFinished, this, &SystemWindow::onInfoLineEditChanged);
    connect(ui->native_lineEdit, &QLineEdit::editingFinished, this, &SystemWindow::onInfoLineEditChanged);
    connect(ui->dorm_lineEdit, &QLineEdit::editingFinished, this, &SystemWindow::onInfoLineEditChanged);
    connect(ui->class_lineEdit, &QLineEdit::editingFinished, this, &SystemWindow::onInfoLineEditChanged);
    connect(ui->birthday_lineEdit, &QLineEdit::editingFinished, this, &SystemWindow::onInfoLineEditChanged);
    connect(ui->gender_combobox, QOverload<int>::of(&QComboBox::currentIndexChanged), this, &SystemWindow::onInfoLineEditChanged);

    // 连接出勤安排按钮的信号与槽
    connect(ui->monday_up_NJH_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->monday_up_NJH_pushButton);});
    connect(ui->monday_up_DXY_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->monday_up_DXY_pushButton);});
    connect(ui->tuesday_up_NJH_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->tuesday_up_NJH_pushButton);});
    connect(ui->tuesday_up_DXY_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->tuesday_up_DXY_pushButton);});
    connect(ui->wednesday_up_NJH_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->wednesday_up_NJH_pushButton);});
    connect(ui->wednesday_up_DXY_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->wednesday_up_DXY_pushButton);});
    connect(ui->thursday_up_NJH_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->thursday_up_NJH_pushButton);});
    connect(ui->thursday_up_DXY_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->thursday_up_DXY_pushButton);});
    connect(ui->friday_up_NJH_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->friday_up_NJH_pushButton);});
    connect(ui->friday_up_DXY_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->friday_up_DXY_pushButton);});
    connect(ui->monday_down_NJH_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->monday_down_NJH_pushButton);});
    connect(ui->monday_down_DXY_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->monday_down_DXY_pushButton);});
    connect(ui->tuesday_down_NJH_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->tuesday_down_NJH_pushButton);});
    connect(ui->tuesday_down_DXY_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->tuesday_down_DXY_pushButton);});
    connect(ui->wednesday_down_NJH_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->wednesday_down_NJH_pushButton);});
    connect(ui->wednesday_down_DXY_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->wednesday_down_DXY_pushButton);});
    connect(ui->thursday_down_NJH_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->thursday_down_NJH_pushButton);});
    connect(ui->thursday_down_DXY_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->thursday_down_DXY_pushButton);});
    connect(ui->friday_down_NJH_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->friday_down_NJH_pushButton);});
    connect(ui->friday_down_DXY_pushButton, &QPushButton::clicked, [this](bool) {onAttendanceButtonClicked(ui->friday_down_DXY_pushButton);});
    // 连接是否执勤按钮的信号与槽
    connect(ui->isWork_pushButton, &QPushButton::clicked, this, &SystemWindow::onIsWorkPushButtonClicked);
}
SystemWindow::~SystemWindow()
{
    delete ui;
}
//关闭窗口事件
void SystemWindow::closeEvent(QCloseEvent *event)
{
    // 弹出提示窗口
    QMessageBox::StandardButton reply = QMessageBox::question(this, "关闭系统", "是否关闭系统？", QMessageBox::Yes | QMessageBox::No);
    if (reply == QMessageBox::Yes) {
        // 保存文件
        FlagGroupFileManager::saveToFile(flagGroup, filename);
        // 接受关闭事件
        event->accept();
    } else {
        // 忽略关闭事件，取消关闭行为
        event->ignore();
    }
}

//值周管理界面函数实现
void SystemWindow::onTabulateButtonClicked() {
    // 制表按钮
    if (!manager) {
        bool useTotalTimesRule = ui->times_rule->isChecked();
        SchedulingManager::HandoverRule handoverRule = SchedulingManager::NoRule;
        if (ui->Monday_handover_rule_radioButton->isChecked()) {
            handoverRule = SchedulingManager::MondayHandoverRule;
        } else if (ui->All_handover_rule_radioButton->isChecked()) {
            handoverRule = SchedulingManager::AllHandoverRule;
        }
        manager = new SchedulingManager(flagGroup, useTotalTimesRule, handoverRule);
        connect(manager, &SchedulingManager::schedulingWarning, this, &SystemWindow::handleSchedulingWarning);  // 连接警告信号与发送警告信息的槽函数
        connect(manager, &SchedulingManager::schedulingFinished, [this]() {
            updateTableWidget(*manager); // 制表操作
            updateTextEdit(*manager); // 更新制表结果文本域
            delete manager;
            manager = nullptr;
        });
    }
    manager->schedule();
}
void SystemWindow::updateTableWidget(const SchedulingManager& manager) {
    //制表操作，点击制表按钮后的辅助函数
    const auto& scheduleTable = manager.getScheduleTable();
    // 从周一上午开始，依次处理表格每个时间槽（周一上午、周一下午、周二上午、周二下午…… 周五下午）
    for (int slot = 0; slot < 10; ++slot) {
        int day = slot / 2; // 0~4，分别对应周一至周五
        int halfDay = slot % 2; // 时段，0~1，分别对应升旗与降旗
        for (int location = 0; location < 2; ++location) {
            // 对于每个时间槽，依次处理两个地点。
            int row = halfDay * 2 + location;// 表格对应的行数，0~3，对应表格单元项的第一到第四行（即不包括表头）
            QString cellText;
            for (int position = 0; position < 3; ++position) {
                // 对于每个地点，检查并添加 3 个人员位置的人员姓名。
                if (scheduleTable[slot][location][position]) {
                    cellText += QString::fromStdString(scheduleTable[slot][location][position]->getName()) + " ";
                }
            }
            ui->worksheet->setItem(row, day, new QTableWidgetItem(cellText.trimmed()));
        }
    }

    // 处理完表格后的表格和窗口大小调整功能

    // 表格大小处理
    // 调整表格列宽以适应内容
    ui->worksheet->resizeColumnsToContents();
    // 调整表格行高以适应内容
    ui->worksheet->resizeRowsToContents();

    // 窗口大小处理
    //
    // 窗口宽度计算
    // 计算表格所需的总宽度
    int totalTableWidth = 0;
    QHeaderView* horizontalHeader = ui->worksheet->horizontalHeader();
    for (int col = 0; col < ui->worksheet->columnCount(); ++col) {
        totalTableWidth += horizontalHeader->sectionSize(col);
    }
    // 加上垂直表头的宽度
    totalTableWidth += ui->worksheet->verticalHeader()->width();
    // 计算表格所需的总高度
    int totalTableHeight = 0;
    QHeaderView* verticalHeader = ui->worksheet->verticalHeader();
    for (int row = 0; row < ui->worksheet->rowCount(); ++row) {
        totalTableHeight += verticalHeader->sectionSize(row);
    }
    //
    // 窗口高度计算
    // 加上水平表头的高度
    totalTableHeight += ui->worksheet->horizontalHeader()->height();
    // 获取 task_toolBox 的宽度
    QWidget* taskToolBox = ui->task_all_splitter->widget(0);
    int taskToolBoxWidth = taskToolBox->width();


    // 获取 timesResult 的高度
    QWidget* timesResult = ui->task_worksheet_splitter->widget(1);
    int timesResultHeight = timesResult->height();
    // 获取窗口当前的布局边距
    QMargins margins = layout()->contentsMargins();
    int marginLeft = margins.left();
    int marginRight = margins.right();
    int marginTop = margins.top();
    int marginBottom = margins.bottom();
    int extraWidth = 80;//补足，人工修改的宽度
    int extraHeight = 10;//补足，人工修改的高度
    // 计算新的宽度和高度
    int newWindowWidth = totalTableWidth + taskToolBoxWidth + marginLeft + marginRight + extraWidth;
    int newWindowHeight = totalTableHeight + timesResultHeight + marginTop + marginBottom + extraHeight;
    // // 调整 task_worksheet_splitter 中 QTableWidget 和 timesResult 的大小
    // QList<int> taskWorksheetSizes;
    // taskWorksheetSizes << totalTableHeight << timesResultHeight;
    // ui->task_worksheet_splitter->setSizes(taskWorksheetSizes);
    // // 调整 task_all_splitter 中 task_toolBox 和 task_worksheet_splitter 的大小
    // QList<int> taskAllSizes;
    // taskAllSizes << taskToolBoxWidth << totalTableWidth;
    // ui->task_all_splitter->setSizes(taskAllSizes);

    // 获取当前窗口的大小
    QSize currentWindowSize = this->size();
    int currentWidth = currentWindowSize.width();
    int currentHeight = currentWindowSize.height();
    // 只放大不缩小
    if (newWindowWidth > currentWidth || newWindowHeight > currentHeight) {
        newWindowWidth = qMax(newWindowWidth, currentWidth);
        newWindowHeight = qMax(newWindowHeight, currentHeight);
        // 调整窗口大小
        resize(newWindowWidth, newWindowHeight);
    }

}
void SystemWindow::handleSchedulingWarning(const QString& warningMessage)
{
    // 排表过程出现无法选出合适人选时的情况，保存警告信息
    warningMessages += warningMessage + "\n";
}

void SystemWindow::updateTextEdit(const SchedulingManager& manager) {
    // 制表结果文本域更新
    QString resultText;
    const auto& availableMembers = manager.getAvailableMembers();
    for (const auto& member : availableMembers) {
        resultText += QString::fromStdString(member->getName()) + " 的工作次数: " + QString::number(member->getTimes()) +
                      " 总工作次数: " + QString::number(member->getAll_times()) + "\n";
    }
    // 拼接警告信息和排班结果文本
    QString finalText = warningMessages + resultText;
    // 设置最终文本到文本编辑框
    ui->timesResult->setPlainText(finalText);
    // 清空警告信息，以便下次排表使用
    warningMessages.clear();
}
void SystemWindow::onClearButtonClicked() {
    //清空表格按钮
    ui->worksheet->clearContents();
    ui->timesResult->clear();
}
void SystemWindow::onResetButtonClicked() {
    //重置队员执勤次数按钮
    for (int i = 1; i < 5; ++i) {
        auto& allMembers = flagGroup.getGroupMembers(i);
        for (auto& member : allMembers) {
            member.setAll_times(0);
        }
    }
    // 在 QTextEdit 中清空并输出提示语句
    ui->timesResult->clear();
    ui->timesResult->append("所有队员的执勤总次数已成功归零");
}
// 导出表格颜色转换辅助函数
int rgbToBgr(const QColor& color) {
    return (color.blue() << 16) | (color.green() << 8) | color.red();
}
void SystemWindow::onExportButtonClicked()
{
    // 导出表格按钮点击事件
    // 获取保存文件路径
    QString filePath = QFileDialog::getSaveFileName(this, "导出表格", "第X周升降旗.xlsx", "Excel 文件 (*.xlsx)");
    // 检查文件路径
    if (!filePath.isEmpty()) {
        // 启动 Excel 应用程序
        QAxObject *excel = new QAxObject("Excel.Application");
        if (excel) {
            // 设置 Excel 应用程序不可见
            excel->dynamicCall("SetVisible(bool)", false);
            // 获取 Excel 应用程序的工作簿集合
            QAxObject *workbooks = excel->querySubObject("Workbooks");
            // 创建一个新的工作簿
            QAxObject *workbook = workbooks->querySubObject("Add");
            if (workbook) {
                // 获取新工作簿的第一个工作表
                QAxObject *worksheetExcel = workbook->querySubObject("Worksheets(int)", 1);
                // 1. 所有拥有文字的区域（A1~G5矩阵区域）都应该居中对齐
                QAxObject *allRange = worksheetExcel->querySubObject("Range(const QString&)", "A1:G5");
                QAxObject *allAlignment = allRange->querySubObject("HorizontalAlignment");
                if (allAlignment) {
                    allAlignment->dynamicCall("SetValue(int)", -4108); // xlCenter
                    delete allAlignment;
                }
                // 2. 第一行列标题C1~G1区域，文本内容不变，字体格式改为16号黑体，背景填充色改为#FFC000
                QAxObject *headerRange = worksheetExcel->querySubObject("Range(const QString&)", "C1:G1");
                QAxObject *headerFont = headerRange->querySubObject("Font");
                headerFont->dynamicCall("SetName(const QString&)", "黑体");
                headerFont->dynamicCall("SetSize(int)", 16);
                QAxObject *headerInterior = headerRange->querySubObject("Interior");
                QColor Color("#FFC000");
                int BgrColor = rgbToBgr(Color);
                headerInterior->dynamicCall("SetColor(int)", BgrColor);
                // 复制列标题到 Excel
                for (int col = 0; col < ui->worksheet->columnCount(); ++col) {
                    QString headerText = ui->worksheet->horizontalHeaderItem(col)->text();
                    QAxObject *cell = worksheetExcel->querySubObject("Cells(int,int)", 1, col + 3); // 从第一行第三列开始写列标题
                    cell->dynamicCall("SetValue(const QVariant&)", headerText);
                }
                // 3. 第一列A2和A3区域合并，并输入“升旗”文本，字体格式改为16号黑体，背景填充色改为#FFFF00
                // 两者使用的颜色单位不同，不能直接转换需要rgb to bgr的操作
                QAxObject *riseFlagRange = worksheetExcel->querySubObject("Range(const QString&)", "A2:A3");
                riseFlagRange->dynamicCall("Merge()");
                QAxObject *riseFlagCell = worksheetExcel->querySubObject("Cells(int,int)", 2, 1);
                riseFlagCell->dynamicCall("SetValue(const QVariant&)", "升旗");
                QAxObject *riseFlagFont = riseFlagRange->querySubObject("Font");
                riseFlagFont->dynamicCall("SetName(const QString&)", "黑体");
                riseFlagFont->dynamicCall("SetSize(int)", 16);
                QAxObject *riseFlagInterior = riseFlagRange->querySubObject("Interior");
                Color = QColor("#FFF000");
                BgrColor = rgbToBgr(Color);
                riseFlagInterior->dynamicCall("SetColor(int)", BgrColor);
                // 4. 第一列A4和A5区域合并，并输入“降旗”文本，字体格式改为16号黑体，背景填充色改为#FFFF00
                QAxObject *lowerFlagRange = worksheetExcel->querySubObject("Range(const QString&)", "A4:A5");
                lowerFlagRange->dynamicCall("Merge()");
                QAxObject *lowerFlagCell = worksheetExcel->querySubObject("Cells(int,int)", 4, 1);
                lowerFlagCell->dynamicCall("SetValue(const QVariant&)", "降旗");
                QAxObject *lowerFlagFont = lowerFlagRange->querySubObject("Font");
                lowerFlagFont->dynamicCall("SetName(const QString&)", "黑体");
                lowerFlagFont->dynamicCall("SetSize(int)", 16);
                QAxObject *lowerFlagInterior = lowerFlagRange->querySubObject("Interior");
                Color = QColor("#FFF000");
                BgrColor = rgbToBgr(Color);
                lowerFlagInterior->dynamicCall("SetColor(int)", BgrColor);
                // 5. 第二列B2~B5区域行标题，文本内容不变，字体格式改为12号等线，背景填充色改为#FFFF00
                QAxObject *rowHeaderRange = worksheetExcel->querySubObject("Range(const QString&)", "B2:B5");
                QAxObject *rowHeaderFont = rowHeaderRange->querySubObject("Font");
                rowHeaderFont->dynamicCall("SetName(const QString&)", "等线");
                rowHeaderFont->dynamicCall("SetSize(int)", 12);
                QAxObject *rowHeaderInterior = rowHeaderRange->querySubObject("Interior");
                Color = QColor("#FFF000");
                BgrColor = rgbToBgr(Color);
                rowHeaderInterior->dynamicCall("SetColor(int)", BgrColor);
                // 复制行标题到 Excel
                for (int row = 0; row < ui->worksheet->rowCount(); ++row) {
                    QString rowHeaderText = ui->worksheet->verticalHeaderItem(row)->text();
                    QAxObject *cell = worksheetExcel->querySubObject("Cells(int,int)", row + 2, 2); // 从第二行第二列写行标题
                    cell->dynamicCall("SetValue(const QVariant&)", rowHeaderText);
                }
                // 复制表格数据到 Excel，从第二行第三列开始
                for (int row = 0; row < ui->worksheet->rowCount(); ++row) {
                    for (int col = 0; col < ui->worksheet->columnCount(); ++col) {
                        QTableWidgetItem *item = ui->worksheet->item(row, col);
                        if (item) {
                            QAxObject *cell = worksheetExcel->querySubObject("Cells(int,int)", row + 2, col + 3);
                            cell->dynamicCall("SetValue(const QVariant&)", item->text());
                        }
                    }
                }
                // // 6. A1~G5矩阵区域设置粗外侧框线
                QAxObject *allBorders = allRange->querySubObject("Borders");
                if (allBorders) {
                    // 设置边框样式为连续线条
                    allBorders->dynamicCall("LineStyle", 1); // xlContinuous
                    // 设置边框粗细为粗线
                    allBorders->dynamicCall("Weight", 2);    // xlThick
                    delete allBorders;
                }
                //C1~G1区域、B2~B5区域、A2和A3的合并区域、A4和A5的合并区域分别设置为所有框线
                QAxObject *headerBorders = headerRange->querySubObject("Borders");
                headerBorders->dynamicCall("LineStyle", 1); // xlContinuous
                QAxObject *rowHeaderBorders = rowHeaderRange->querySubObject("Borders");
                rowHeaderBorders->dynamicCall("LineStyle", 1); // xlContinuous
                QAxObject *riseFlagBorders = riseFlagRange->querySubObject("Borders");
                riseFlagBorders->dynamicCall("LineStyle", 1); // xlContinuous
                QAxObject *lowerFlagBorders = lowerFlagRange->querySubObject("Borders");
                lowerFlagBorders->dynamicCall("LineStyle", 1); // xlContinuous
                // 7. A列宽110像素，B列宽175像素，C~G列宽350像素，1~5行高全部设置为65像素
                // 两者计量单位不同，需要计算后转换
                QAxObject *columnA = worksheetExcel->querySubObject("Columns(const QString&)", "A");
                columnA->dynamicCall("ColumnWidth", 8);
                QAxObject *columnB = worksheetExcel->querySubObject("Columns(const QString&)", "B");
                columnB->dynamicCall("ColumnWidth", 14);
                QAxObject *columnsCToG = worksheetExcel->querySubObject("Range(const QString&)", "C:G");
                columnsCToG->dynamicCall("ColumnWidth", 28);
                QAxObject *rows1To5 = worksheetExcel->querySubObject("Range(const QString&)", "1:5");
                rows1To5->dynamicCall("RowHeight", 32.5);
                // 调整列宽以适应内容（可根据需要保留或移除）
                // QAxObject *usedRange = worksheetExcel->querySubObject("UsedRange");
                // usedRange->querySubObject("Columns")->dynamicCall("AutoFit()");
                // 保存并关闭 Excel 文件
                // 将工作簿保存到用户指定的路径
                workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(filePath));
                // 关闭工作簿
                workbook->dynamicCall("Close()");
            }
            // 退出 Excel 应用程序
            excel->dynamicCall("Quit()");
            // 释放 Excel 应用程序对象的内存
            delete excel;
            // 显示导出结果提示
            QMessageBox::information(this, "导出成功", "表格已成功导出到指定位置。");
        } else {
            // 显示导出结果提示
            QMessageBox::critical(this, "导出失败", "无法启动 Excel 应用程序，请确保已安装 Excel。");
        }
    }
}
void SystemWindow::onTotalTimesRuleClicked() {
    // 总次数规则按钮，选定或取消选定
    if (manager) {
        manager->setUseTotalTimesRule(ui->times_rule->isChecked());
    }
}

void SystemWindow::onRadioButtonClicked()
{
    // 当单选按钮被点击时，更新交接规则信息
    if (manager) {
        if (ui->Monday_handover_rule_radioButton->isChecked()) {
            manager->setHandoverRule(SchedulingManager::MondayHandoverRule);
        } else if (ui->All_handover_rule_radioButton->isChecked()) {
            manager->setHandoverRule(SchedulingManager::AllHandoverRule);
        } else if (ui->No_handover_rule_radioButton->isChecked()) {
            manager->setHandoverRule(SchedulingManager::NoRule);
        }
    }
}



//队员管理界面函数实现
void SystemWindow::onGroupAddButtonClicked(int groupIndex)
{
    //添加队员按钮点击事件
    bool isChecked = false;
    //判断是哪个组发出的信号，修改新队员的是否值周状态
    switch (groupIndex) {
    case 1: isChecked = ui->group1_iswork_radioButton->isChecked(); break;
    case 2: isChecked = ui->group2_iswork_radioButton->isChecked(); break;
    case 3: isChecked = ui->group3_iswork_radioButton->isChecked(); break;
    case 4: isChecked = ui->group4_iswork_radioButton->isChecked(); break;
    } 
    bool time[4][5];
    //初始化所有执勤时间，默认为全部可以执勤
    for (int i = 0; i < 4; ++i) {
        for (int g = 0; g < 5; ++g) {
            time[i][g] = false;
        }
    }

    // 生成唯一的默认名字
    std::string defaultNameBase = "未命名队员";
    int counter = 1;
    std::string defaultName = defaultNameBase + std::to_string(counter);

    // 获取所有队员的名字
    std::vector<std::string> existingNames;
    for (int i = 1; i <= 4; ++i) {
        const auto& members = flagGroup.getGroupMembers(i);
        for (const auto& member : members) {
            existingNames.push_back(member.getName());
        }
    }

    // 检查名字是否重复，若重复则递增计数器
    while (std::find(existingNames.begin(), existingNames.end(), defaultName) != existingNames.end()) {
        counter++;
        defaultName = defaultNameBase + std::to_string(counter);
    }
    //初始化新队员的基础信息
    Person person(defaultName, false, groupIndex, "", "", "", "", "", "", "", isChecked, time, 0, 0);
    //向flagGroup中添加新队员
    flagGroup.addPersonToGroup(person, groupIndex);
    //更新对应组的ListView组员标签信息
    updateListView(groupIndex);
}
void SystemWindow::onGroupDeleteButtonClicked(int groupIndex)
{
    //删除队员按钮点击事件
    QListView* listView = nullptr;
    //判断是哪个组发出的信号，以及信号情况
    switch (groupIndex) {
    case 1: listView = ui->group1_info_listView; break;
    case 2: listView = ui->group2_info_listView; break;
    case 3: listView = ui->group3_info_listView; break;
    case 4: listView = ui->group4_info_listView; break;
    }
    //如果出现非法组号，退出
    if (!listView) return;
    QModelIndexList selectedIndexes = listView->selectionModel()->selectedIndexes();
    //不存在被选定的队员
    if (selectedIndexes.isEmpty()) return;
    //删除前的确认弹窗
    QMessageBox::StandardButton reply = QMessageBox::question(this, "确认删除", "是否删除该队员？", QMessageBox::Yes | QMessageBox::No);
    //删除对应队员
    if (reply == QMessageBox::Yes) {
        int row = selectedIndexes.first().row();
        const auto& members = flagGroup.getGroupMembers(groupIndex);
        if (static_cast<std::vector<Person>::size_type>(row) < members.size()) {
            const Person& personToRemove = members[row];
            flagGroup.removePersonFromGroup(personToRemove, groupIndex);// 调用 Flag_group 的删除成员方法
            updateListView(groupIndex);//更新对应组的ListView组员标签信息
        }
    }
}
void SystemWindow::onGroupIsWorkRadioButtonClicked(int groupIndex)
{
    //是否值周确认按钮点击事件
    bool isChecked = false;
    //判断是哪个组发出的信号，以及信号情况
    switch (groupIndex) {
    case 1: isChecked = ui->group1_iswork_radioButton->isChecked(); break;
    case 2: isChecked = ui->group2_iswork_radioButton->isChecked(); break;
    case 3: isChecked = ui->group3_iswork_radioButton->isChecked(); break;
    case 4: isChecked = ui->group4_iswork_radioButton->isChecked(); break;
    }
    //获取对应组别所有队员
    const auto& members = flagGroup.getGroupMembers(groupIndex);
    //设置对应组别所有队员isWork属性，选中设为1，取消选中设为0
    for (auto& member : const_cast<std::vector<Person>&>(members)) {
        member.setIsWork(isChecked);
    }
}
void SystemWindow::onListViewItemClicked(const QModelIndex &index, int groupIndex)
{
    //队员标签点击事件
    //显示选中队员的信息
    Person* person = getSelectedPerson(groupIndex, index);//捕捉被点击的队员是谁
    if (person) {
        currentSelectedPerson = person;
        isShowingInfo = true; // 设置标志位为展示信息状态
        showMemberInfo(*person);//显示基础信息
        updateAttendanceButtons(*person);//显示执勤信息
        isShowingInfo = false; // 恢复标志位
    }
}
Person* SystemWindow::getSelectedPerson(int groupIndex, const QModelIndex &index)
{
    //捕捉被选中的标签是哪个队员
    const auto& members = flagGroup.getGroupMembers(groupIndex);
    if (index.isValid() && static_cast<std::vector<Person>::size_type>(index.row()) < members.size()){
        return const_cast<Person*>(&members[index.row()]);
    }
    return nullptr;
}
void SystemWindow::onInfoLineEditChanged()
{
    // 信息修改后更新 Flag_group 中队员的信息,不包括点击队员标签时显示队员信息时造成的修改
    if (currentSelectedPerson && !isShowingInfo) {
        updatePersonInfo(*currentSelectedPerson);
    }
}
void SystemWindow::onGroupComboBoxChanged(int newGroupIndex)
{
    // 独立的修改组别函数
    if (currentSelectedPerson && !isShowingInfo) {
        int oldGroupIndex = currentSelectedPerson->getGroup();//值为1~4
        if((oldGroupIndex - 1) != newGroupIndex)//规避并未修改组别引发多余操作
        {
            newGroupIndex += 1; // combobox 索引从 0 开始，组索引从 1 开始
            currentSelectedPerson->setGroup(newGroupIndex);// 更新队员的组别信息
            bool isChecked = false;
            switch (newGroupIndex) {
            case 1: isChecked = ui->group1_iswork_radioButton->isChecked(); break;
            case 2: isChecked = ui->group2_iswork_radioButton->isChecked(); break;
            case 3: isChecked = ui->group3_iswork_radioButton->isChecked(); break;
            case 4: isChecked = ui->group4_iswork_radioButton->isChecked(); break;
            }
            currentSelectedPerson->setIsWork(isChecked);
            flagGroup.addPersonToGroup(*currentSelectedPerson, newGroupIndex);// 将队员添加到新组中
            updateListView(newGroupIndex);// 更新新组组别信息
            Person* person = flagGroup.findPersonInGroup(*currentSelectedPerson, newGroupIndex);//创建一个新的队员指针跟踪新创建的队员
            flagGroup.removePersonFromGroup(*currentSelectedPerson, oldGroupIndex);// 从旧组中删除队员
            updateListView(oldGroupIndex);// 更新旧组组别信息
            currentSelectedPerson = person;// 更新当前选中的队员指针
        }
    }
}
void SystemWindow::updateListView(int groupIndex)
{
    // 更新队员标签界面
    QListView* listView = nullptr;
    switch (groupIndex) {
    case 1: listView = ui->group1_info_listView; break;
    case 2: listView = ui->group2_info_listView; break;
    case 3: listView = ui->group3_info_listView; break;
    case 4: listView = ui->group4_info_listView; break;
    }
    if (!listView) return;
    const auto& members = flagGroup.getGroupMembers(groupIndex);
    QStringList memberNames;
    for (const auto& member : members) {
        memberNames << QString::fromStdString(member.getName());
    }
    // 使用 QStringListModel 替代 QStandardItemModel
    QStringListModel *model = new QStringListModel();//创建一个 QStringListModel 对象 model，它是 QAbstractItemModel 的子类，专门用于处理字符串列表数据
    model->setStringList(memberNames);//调用 model 的 setStringList 方法，将 memberNames 列表中的数据设置到模型中。
    listView->setModel(model);//调用 listView 的 setModel 方法，将 model 设置为 listView 的数据模型。这样，listView 就会显示 model 中的数据，即指定组别的队员姓名。
}
void SystemWindow::showMemberInfo(const Person &person)
{
    // 根据选中的队员向UI中展示队员基础信息
    ui->name_lineEdit->setText(QString::fromStdString(person.getName()));
    ui->phone_lineEdit->setText(QString::fromStdString(person.getPhone_number()));
    ui->nativePlace_lineEdit->setText(QString::fromStdString(person.getNative_place()));
    ui->school_lineEdit->setText(QString::fromStdString(person.getSchool()));
    ui->native_lineEdit->setText(QString::fromStdString(person.getNative()));
    ui->dorm_lineEdit->setText(QString::fromStdString(person.getDorm()));
    ui->class_lineEdit->setText(QString::fromStdString(person.getClassname()));
    ui->birthday_lineEdit->setText(QString::fromStdString(person.getBirthday()));
    ui->gender_combobox->setCurrentText(person.getGender() ? "女" : "男");
    ui->group_combobox->setCurrentIndex(person.getGroup() - 1);//group_combobox默认以0开始，与组名最小为1相违背
}
void SystemWindow::updatePersonInfo(const Person &person)
{
    // 仅更新基础信息部分，执勤安排不调整
    // 从 UI 中获取更新后的信息，修改flag_group中队员信息
    QString name = ui->name_lineEdit->text();
    QString phone = ui->phone_lineEdit->text();
    QString nativePlace = ui->nativePlace_lineEdit->text();
    QString school = ui->school_lineEdit->text();
    QString native = ui->native_lineEdit->text();
    QString dorm = ui->dorm_lineEdit->text();
    QString classname = ui->class_lineEdit->text();
    QString birthday = ui->birthday_lineEdit->text();
    bool gender = ui->gender_combobox->currentText() == "女";
    // 创建新的 Person 对象
    bool time[4][5];
    //time 数组保持不变
    for (int i = 0; i < 4; ++i) {
        for (int j = 0; j < 5; ++j) {
            time[i][j] = person.getTime(i + 1, j + 1);
        }
    }
    Person newPerson(name.toStdString(), gender, person.getGroup(), phone.toStdString(),
                     nativePlace.toStdString(), native.toStdString(), dorm.toStdString(),
                     school.toStdString(), classname.toStdString(), birthday.toStdString(), person.getIsWork(),
                     time, person.getTimes(), person.getAll_times());
    // 更新 flagGroup 中对应队员的信息
    flagGroup.modifyPersonInGroup(person, newPerson, person.getGroup());
    // 更新对应组的 ListView 显示
    updateListView(person.getGroup());

}
void SystemWindow::updateAttendanceButtons(const Person &person)
{
    // 根据队员的time数组调整按钮显示的状态
    // 周一升旗
    ui->monday_up_NJH_pushButton->setChecked(person.getTime(1, 1));
    ui->monday_up_DXY_pushButton->setChecked(person.getTime(2, 1));
    // 周二升旗
    ui->tuesday_up_NJH_pushButton->setChecked(person.getTime(1, 2));
    ui->tuesday_up_DXY_pushButton->setChecked(person.getTime(2, 2));
    // 周三升旗
    ui->wednesday_up_NJH_pushButton->setChecked(person.getTime(1, 3));
    ui->wednesday_up_DXY_pushButton->setChecked(person.getTime(2, 3));
    // 周四升旗
    ui->thursday_up_NJH_pushButton->setChecked(person.getTime(1, 4));
    ui->thursday_up_DXY_pushButton->setChecked(person.getTime(2, 4));
    // 周五升旗
    ui->friday_up_NJH_pushButton->setChecked(person.getTime(1, 5));
    ui->friday_up_DXY_pushButton->setChecked(person.getTime(2, 5));
    // 周一降旗
    ui->monday_down_NJH_pushButton->setChecked(person.getTime(3, 1));
    ui->monday_down_DXY_pushButton->setChecked(person.getTime(4, 1));
    // 周二降旗
    ui->tuesday_down_NJH_pushButton->setChecked(person.getTime(3, 2));
    ui->tuesday_down_DXY_pushButton->setChecked(person.getTime(4, 2));
    // 周三降旗
    ui->wednesday_down_NJH_pushButton->setChecked(person.getTime(3, 3));
    ui->wednesday_down_DXY_pushButton->setChecked(person.getTime(4, 3));
    // 周四降旗
    ui->thursday_down_NJH_pushButton->setChecked(person.getTime(3, 4));
    ui->thursday_down_DXY_pushButton->setChecked(person.getTime(4, 4));
    // 周五降旗
    ui->friday_down_NJH_pushButton->setChecked(person.getTime(3, 5));
    ui->friday_down_DXY_pushButton->setChecked(person.getTime(4, 5));
}
void SystemWindow::onAttendanceButtonClicked(QAbstractButton *button)
{
    // 执勤按钮点击事件
    // 根据按钮修改time数组信息
    Person* currentPerson = currentSelectedPerson;
    if (currentPerson) {
        // 定义按钮到 (row, column) 的映射,使用映射表来存储按钮和对应的 setTime 方法所需的行、列参数，这样可以避免大量重复的条件判断。
        static QMap<QAbstractButton*, std::pair<int, int>> buttonToTimeMap = {
            {ui->monday_up_NJH_pushButton, {1, 1}},
            {ui->monday_up_DXY_pushButton, {2, 1}},
            {ui->tuesday_up_NJH_pushButton, {1, 2}},
            {ui->tuesday_up_DXY_pushButton, {2, 2}},
            {ui->wednesday_up_NJH_pushButton, {1, 3}},
            {ui->wednesday_up_DXY_pushButton, {2, 3}},
            {ui->thursday_up_NJH_pushButton, {1, 4}},
            {ui->thursday_up_DXY_pushButton, {2, 4}},
            {ui->friday_up_NJH_pushButton, {1, 5}},
            {ui->friday_up_DXY_pushButton, {2, 5}},
            {ui->monday_down_NJH_pushButton, {3, 1}},
            {ui->monday_down_DXY_pushButton, {4, 1}},
            {ui->tuesday_down_NJH_pushButton, {3, 2}},
            {ui->tuesday_down_DXY_pushButton, {4, 2}},
            {ui->wednesday_down_NJH_pushButton, {3, 3}},
            {ui->wednesday_down_DXY_pushButton, {4, 3}},
            {ui->thursday_down_NJH_pushButton, {3, 4}},
            {ui->thursday_down_DXY_pushButton, {4, 4}},
            {ui->friday_down_NJH_pushButton, {3, 5}},
            {ui->friday_down_DXY_pushButton, {4, 5}}
        };
        // 查找按钮对应的 (row, column)
        auto it = buttonToTimeMap.find(button);
        if (it != buttonToTimeMap.end()) {
            int row = it.value().first;
            int column = it.value().second;
            currentPerson->setTime(row, column, button->isChecked());
        }
    }
}
void SystemWindow::onAllSelectButtonClicked()
{
    //全选按钮点击事件
    QAbstractButton* senderButton = qobject_cast<QAbstractButton*>(sender());
    if (!senderButton || !currentSelectedPerson) return;
    // 获取当前点击的“全选”按钮所在的 groupBox
    QGroupBox* parentGroupBox = qobject_cast<QGroupBox*>(senderButton->parent());
    if (!parentGroupBox) return;
    // 获取 groupBox 中的其他两个按钮
    QList<QAbstractButton*> childButtons = parentGroupBox->findChildren<QAbstractButton*>();
    for (QAbstractButton* button : childButtons) {
        if (button != senderButton) {
            button->setChecked(true);
        }
    }
    // 更新对应 time 数组
    static QMap<QAbstractButton*, std::pair<int, int>> buttonToTimeMap = {
        {ui->monday_up_NJH_pushButton, {1, 1}},
        {ui->monday_up_DXY_pushButton, {2, 1}},
        {ui->tuesday_up_NJH_pushButton, {1, 2}},
        {ui->tuesday_up_DXY_pushButton, {2, 2}},
        {ui->wednesday_up_NJH_pushButton, {1, 3}},
        {ui->wednesday_up_DXY_pushButton, {2, 3}},
        {ui->thursday_up_NJH_pushButton, {1, 4}},
        {ui->thursday_up_DXY_pushButton, {2, 4}},
        {ui->friday_up_NJH_pushButton, {1, 5}},
        {ui->friday_up_DXY_pushButton, {2, 5}},
        {ui->monday_down_NJH_pushButton, {3, 1}},
        {ui->monday_down_DXY_pushButton, {4, 1}},
        {ui->tuesday_down_NJH_pushButton, {3, 2}},
        {ui->tuesday_down_DXY_pushButton, {4, 2}},
        {ui->wednesday_down_NJH_pushButton, {3, 3}},
        {ui->wednesday_down_DXY_pushButton, {4, 3}},
        {ui->thursday_down_NJH_pushButton, {3, 4}},
        {ui->thursday_down_DXY_pushButton, {4, 4}},
        {ui->friday_down_NJH_pushButton, {3, 5}},
        {ui->friday_down_DXY_pushButton, {4, 5}}
    };
    for (QAbstractButton* button : childButtons) {
        if (button != senderButton) {
            auto it = buttonToTimeMap.find(button);
            if (it != buttonToTimeMap.end()) {
                int row = it.value().first;
                int column = it.value().second;
                currentSelectedPerson->setTime(row, column, true);
            }
        }
    }
}
void SystemWindow::onIsWorkPushButtonClicked()
{
    // 全选/清空按钮点击事件
    // 切换所有出勤按钮的选中状态，并更新 Person 的 time 数组。
    static bool isAllChecked = false;// 一次创建，全局生命周期，第一次点击实现全选功能
    isAllChecked = !isAllChecked;
    // 获取 availableTime_groupBox 中的所有按钮
    QList<QAbstractButton*> attendanceButtons = ui->availableTime_groupBox->findChildren<QAbstractButton*>();
    // 遍历所有按钮，设置选中状态
    for (QAbstractButton* button : attendanceButtons) {
        // 排除以 all_pushButton 结尾的全选按钮
        if (!button->objectName().endsWith("all_pushButton")) {
            button->setChecked(isAllChecked);
        }
    }
    // 调用当前选中的队员信息
    Person* person = currentSelectedPerson;
    if (person) {
        // 根据 isAllChecked 更新 time 数组
        bool newTime[4][5];
        for (int i = 0; i < 4; ++i) {
            for (int j = 0; j < 5; ++j) {
                newTime[i][j] = isAllChecked;
            }
        }
        person->setTime(newTime);
    }
}
