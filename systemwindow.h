// systemwindow.h头文件
// 功能说明：系统的窗口类。在对应源文件下将实现全部的功能

#ifndef SYSTEMWINDOW_H
#define SYSTEMWINDOW_H

#include <QMainWindow>
#include "dataFunction.h"
#include "qabstractbutton.h"
QT_BEGIN_NAMESPACE
namespace Ui {
class SystemWindow;
}
QT_END_NAMESPACE
class SystemWindow : public QMainWindow
{
    Q_OBJECT
public:
    SystemWindow(QWidget *parent = nullptr);
    ~SystemWindow();
protected:
    void closeEvent(QCloseEvent *event) override; // 重写 closeEvent 函数，自定义窗口关闭事件
private slots:
    // 值周管理界面槽函数
    // 表格管理
    void onTabulateButtonClicked(); // 排表按钮点击事件
    void onClearButtonClicked(); // 清空表格按钮点击事件
    void onResetButtonClicked(); // 重置队员执勤总次数按钮点击事件
    void onExportButtonClicked(); // 导出表格按钮点击事件
    // 规则管理
    void onTotalTimesRuleClicked(); // 总次数规则按钮点击事件
    void onRadioButtonClicked(); // 交接工作单选按钮点击事件
    // 制表警告
    void handleSchedulingWarning(const QString& warningMessage); // 排表过程中发送警告信息的对应处理函数

    // 队员管理界面槽函数
    // 组员管理工具栏
    void onGroupAddButtonClicked(int groupIndex); // 添加组员按钮点击事件
    void onGroupDeleteButtonClicked(int groupIndex); // 删除组员按钮点击事件
    void onGroupIsWorkRadioButtonClicked(int groupIndex); // 全组是否执勤按钮点击事件
    void onListViewItemClicked(const QModelIndex &index, int groupIndex); // 队员标签点击事件
    // 队员基础信息栏
    void onInfoLineEditChanged(); // 队员基础信息文本框修改事件
    void onGroupComboBoxChanged(int newGroupIndex); // 队员组别信息修改事件
    // 队员事件安排栏
    void onAttendanceButtonClicked(QAbstractButton *button); // 事件安排表中执勤按钮点击事件
    void onIsWorkPushButtonClicked(); // 总全选/清空按钮点击事件
    void onAllSelectButtonClicked(); // 全选按钮点击事件
private:
    Ui::SystemWindow *ui; // ui界面指针
    SchedulingManager *manager; // 国旗班制表管理器指针
    Flag_group flagGroup; // 国旗班成员容器变量
    Person* currentSelectedPerson = nullptr; // 保存当前用户选中的队员标签指针
    bool isShowingInfo = false; // 新增标志位，用于区分展示信息造成的文本框信息修改和用户主动填写造成的信息修改
    QString warningMessages; // 新增成员变量，用于保存警告信息
    QString filename = "./data/data.txt"; // 保存队员信息的文件名

    // 值周管理操作函数
    void updateTableWidget(const SchedulingManager& manager); // 制表操作，点击制表按钮后的辅助函数
    void updateTextEdit(const SchedulingManager& manager); // 制表结果在文本域中更新，点击制表按钮后的辅助函数
    // 队员管理操作函数
    void updateListView(int groupIndex); // 更新队员标签界面
    Person* getSelectedPerson(int groupIndex, const QModelIndex &index); // 捕捉被选中的标签是哪个队员，队员标签点击后的辅助函数
    void showMemberInfo(const Person &person); // 根据选中的队员向UI中展示队员基础信息
    void updatePersonInfo(const Person &person); // 从UI中获取更新后的信息，修改flag_group中队员信息，仅更新基础信息部分，执勤安排不调整（根据程序实际设计，队员组别信息修改不在该函数进行）。
    void updateAttendanceButtons(const Person &person); // 根据队员的time数组调整按钮显示的状态
};
#endif // SYSTEMWINDOW_H
