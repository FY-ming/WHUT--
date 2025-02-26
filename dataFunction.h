// dataFunction.h头文件
// 功能说明：主要用于提供执勤工作表的相关数据处理函数

// 有关表格信息的额外说明
// scheduleTable索引全都参照下面的表格结构
// 表格结构大致如下
//          ||        周一         ||        周二         ||        周三        ||        周四         ||        周五         ||
//      |南鉴湖升旗 ||slot:0   location:0 || slot:2   location:0||slot:4   location:0 ||slot:6   location:0 ||slot:8   location:0 ||
// 升旗  --------------------------------------------------------------------------------------------------------------------------
//      |东西院升旗 ||slot:0   location:1 ||slot:2   location:1 ||slot:4   location:1 ||slot:6   location:1 ||slot:8   location:1 ||
// -------------------------------------------------------------------------------------------------------------------------------
//      |南鉴湖降旗 ||slot:1   location:0 ||slot:3   location:0 ||slot:5   location:0 ||slot:7   location:0 ||slot:9   location:0 ||
// 降旗  --------------------------------------------------------------------------------------------------------------------------
//      |东西院降旗 ||slot:1   location:1 ||slot:3   location:1 ||slot:5   location:1 ||slot:7   location:1 ||slot:9   location:1 ||

#pragma once

#include <QObject>
#include <QDebug>
#include <vector>
#include <algorithm>
#include <random>
#include <string>
#include "Person.h"
#include "Flag_group.h"


// SchedulingManager 类定义，执勤工作表
class SchedulingManager : public QObject
{
    Q_OBJECT // QObject宏定义

signals:
    void schedulingWarning(const QString& warningMessage); // 某一时间段无法安排队员执勤时的警告信号
    void schedulingFinished(); // 排表完成后的提示信号

public:
    // 针对南鉴湖交接规则的枚举成员
    enum HandoverRule {
        NoRule, // 不采用交接规则
        MondayHandoverRule, // 仅周二的南鉴湖升旗采用交接规则
        AllHandoverRule // 全周（周二至周五）南鉴湖升旗采用交接规则
    };
    // 构造函数
    SchedulingManager(const Flag_group& flagGroup, bool useTotalTimesRule = false, HandoverRule handoverRule = NoRule)
        : flagGroup(flagGroup), useTotalTimesRule(useTotalTimesRule), handoverRule(handoverRule) {
        initializeAvailableMembers();// 通过队员的isWork的信息统计参加排班的人
    }
    // 部署工作表基础准备资源，排班操作的入口
    void schedule() {
        const int totalSlots = 10;// 一周10个工作时间段,升旗时间对应0 2 4 6 8
        const int locationsPerSlot = 2;// 两个工作地点（0:南鉴湖、1:东西院）
        const int peoplePerLocation = 3;// 一个工作地点的三名执勤队员
        for (auto& member : availableMembers) {
            // 重置每个参加排班的队员本周的工作次数：0
            member->setTimes(0);
        }

        // 随机数生成装置，生成高质量随机数种子，理论上不可能重复
        std::random_device rd;
        // 创建一个 std::mt19937 类型的随机数引擎 g，并使用 rd() 生成的随机数种子对其进行初始化。
        // std::mt19937 是一个基于梅森旋转算法的伪随机数生成器，能够生成高质量的随机数序列。
        // 理论上存在重复的可能，但运算周期极长，只要随机数种子不同，理论不会重复。
        // 因为是伪随机数，所以种子一样，结果一样
        std::mt19937 g(rd());
        // 调用 std::shuffle 函数，将 availableMembers 向量中的元素顺序随机打乱。
        // std::shuffle 函数接受三个参数：容器的起始迭代器、容器的结束迭代器以及随机数引擎。
        // 借助 std::shuffle 函数，能够将 availableMembers 向量中的队员指针顺序随机打乱。
        // 在分配剩余工作量时，每个队员都有相同的概率获得额外的工作机会，避免了因队员在列表中的初始顺序而导致的不公平现象。
        // 若不使用 std::shuffle 对 availableMembers 进行随机打乱，那么每次剩余工作量都会优先分配给列表前面的队员。
        // 长期下来，这会造成队员之间的工作量不均衡，前面的队员工作次数会明显多于后面的队员。
        std::shuffle(availableMembers.begin(), availableMembers.end(), g);


        // 排班！
        // 对 scheduleTable 进行初始化，它是一个三维向量，用于存储排班结果。
        // totalSlots：表示一周内的总工作时间段数量。在当前的排班规则下，一周工作 5 天，每天分上午和下午两个时间段，所以 totalSlots 为 10。
        // locationsPerSlot：每个工作时间段内的工作地点数量，这里是 2 个（“NJH” 和 “DXY”）。
        // peoplePerLocation：每个工作地点需要的工作人员数量，这里是 3 人。
        // nullptr：初始时，每个排班位置都设置为 nullptr，表示尚未安排人员。
        scheduleTable.resize(totalSlots, std::vector<std::vector<Person*>>(locationsPerSlot, std::vector<Person*>(peoplePerLocation, nullptr)));
        for (int slot = 0; slot < totalSlots; ++slot) {
            //外层循环遍历工作时间段
            int day = slot / 2 + 1;//值为1~5。表示星期
            int halfDay = slot % 2;//值为0~1。0:上午，1：下午
            for (int location = 0; location < locationsPerSlot; ++location) {
                // 中层循环遍历工作地点
                int timeRow = halfDay * 2 + location + 1;//location=0~1,timeRow=1~4，分别表示NJH升旗，DXY升旗，NJH降旗，DXY降旗
                for (int position = 0; position < peoplePerLocation; ++position) {
                    //内层循环遍历工作岗位
                    Person* selectedPerson = selectPerson(slot, timeRow, location, day);//选择合适队员
                    if (selectedPerson) {
                        // 如果找到合适队员，加入工作表格scheduleTable中
                        scheduleTable[slot][location][position] = selectedPerson;
                        selectedPerson->setTimes(selectedPerson->getTimes() + 1);
                        selectedPerson->setAll_times(selectedPerson->getAll_times() + 1);
                    }
                }
            }
        }
        // 发出排班完成信号
        emit schedulingFinished();
    }

    // 成员变量的get与set函数声明
    bool getUseTotalTimesRule() const;
    void setUseTotalTimesRule(bool newUseTotalTimesRule);
    const Flag_group &getFlagGroup() const;
    std::vector<Person *> getAvailableMembers() const;
    void setAvailableMembers(const std::vector<Person *> &newAvailableMembers);
    std::vector<std::vector<std::vector<Person *> > > getScheduleTable() const;
    void setScheduleTable(const std::vector<std::vector<std::vector<Person *> > > &newScheduleTable);
    HandoverRule getHandoverRule() const;
    void setHandoverRule(HandoverRule newHandoverRule);

private:
    const Flag_group& flagGroup; // 国旗班容器，保存队员信息
    bool useTotalTimesRule; // 规则标签，判断是否使用总次数规则
    HandoverRule handoverRule; // 规则标签，判断是否使用交接规则
    std::unordered_map<std::string, int> warningCount; // 键值对容器，用于记录交接规则失败警告信息出现的次数
    std::vector<Person*> availableMembers; // 容器，保存参加排班的队员
    std::vector<std::vector<std::vector<Person*>>> scheduleTable; // 工作表格

    void initializeAvailableMembers() {
        // 初始化辅助函数
        // 通过队员的isWork的信息统计参加排班的人
        for (int group = 1; group <= 4; ++group) {
            const auto& members = flagGroup.getGroupMembers(group);
            for (const auto& member : members) {
                if (member.getIsWork()) {
                    availableMembers.push_back(const_cast<Person*>(&member));
                }
            }
        }
    }
    Person* selectPerson(int slot, int timeRow,int location, int day) {
        // 制表辅助函数
        // 选择合适的可工作队员
        // slot=0~9，表示10个时间段（周一上午、周一下午、周二上午、周二下午…… 周五下午）
        // timeRow=1~4，表格行数，分别表示NJH升旗，DXY升旗，NJH降旗，DXY降旗
        // location=0~1，工作地点，分别表示南鉴湖，东西院
        // day = 1~5, 工作的时间，对应周一至周五
        if (useTotalTimesRule) {
            // 采用总次数排班
            // 当 useTotalTimesRule 为 true 时，使用 std::sort 函数对 availableMembers 列表进行排序
            // 排序依据是人员的总工作次数（通过 getAll_times() 方法获取），按照总工作次数从小到大排序。
            // 这样做的目的是优先安排总工作次数较少的人员，使得人员的总工作量更加平均。
            std::sort(availableMembers.begin(), availableMembers.end(), [](Person* a, Person* b) {
                return a->getAll_times() < b->getAll_times();
            });
        } else {
            // 普通排班
            // 当 useTotalTimesRule 为 false 时，同样使用 std::sort 函数对 availableMembers 列表进行排序
            // 但排序依据是人员本周的工作次数（通过 getTimes() 方法获取），按照本周工作次数从小到大排序。
            // 这样可以优先安排本周工作次数较少的人员，保证本周内人员工作量的平均分配。
            std::sort(availableMembers.begin(), availableMembers.end(), [](Person* a, Person* b) {
                return a->getTimes() < b->getTimes();
            });
        }
        // 人员筛选
        // 交接规则的人员筛选
        // 如果未选择交接规则，此筛选与下面的普通筛选无异
        // isPersonSatisfyHandoverRule函数：判断person是否符合交接规则。
        // getTime函数：检查person对应时间是否有有空
        // isPersonBusy函数：检查person是否在该时间段已经安排了工作
        for (auto person : availableMembers) {
            if (isPersonSatisfyHandoverRule(person, slot, location) && person->getTime(timeRow, day)  && !isPersonBusy(person, slot)){
                return person;
            }
        }

        // 警告信息临时变量
        std::string days[] = { "周一", "周二", "周三", "周四", "周五" };
        std::string halves[] = { "上午", "下午" };
        std::string locations[] = { "NJH", "DXY" };
        int dayIndex = (slot / 2);
        int halfDayIndex = (slot % 2);

        // 如果无法完成交接规则，将发出警报，放弃交接规则，重新选人
        // 考虑到每次任务有三名队员，交接规则原则上最少只需要有一个队员完成交接即可，所以需要当一次任务的三个队员都不符合交接规则时才发送警告信息
        // 生成无法完成交接规则的警告信息
        std::string warning = "警告：在 " + days[dayIndex] + " " + halves[halfDayIndex] + " " + locations[location] + " 无法完成交接规则。";
        // 增加该警告信息的计数
        warningCount[warning]++;
        // 当警告信息出现三次时才发送
        if (warningCount[warning] == 3) {
            emit schedulingWarning(QString::fromStdString(warning));
        }

        // 普通筛选
        // 当用户采用交接规则但无法找出合适的队员时，将放弃交接规则，采用普通筛选，找到可执勤队员
        // getTime函数：检查person对应时间是否有有空
        // isPersonBusy函数：检查person是否在该时间段已经安排了工作
        for (auto person : availableMembers) {
            if (person->getTime(timeRow, day) && !isPersonBusy(person, slot)) {
                return person;
            }
        }
        // 普通筛选仍无法找到合适队员，系统将发送警告信息
        warning = "警告：在 " + days[dayIndex] + " " + halves[halfDayIndex] + " " + locations[location] + " 无法选出合适的人员进行排班。";
        emit schedulingWarning(QString::fromStdString(warning));
        return nullptr;
    }
    // 判断是否已经在同一时间段安排了工作
    bool isPersonBusy(Person* person, int slot) const {
        // 将对应时间段slot的所有位置都遍历一遍，查看是否已经存在该队员person
        for (int location = 0; location < 2; ++location) {
            for (int position = 0; position < 3; ++position) {
                if (scheduleTable[slot][location][position] == person) {
                    return true;
                }
            }
        }
        return false;
    }

    // 判断人员是否满足交接规则
    bool isPersonSatisfyHandoverRule(Person* person, int slot, int location) {
        switch (handoverRule) {
            case MondayHandoverRule: // 仅周二的南鉴湖升旗采用交接规则
            {
                if (slot == 2 && location == 0) { // 对应表格一行二列，周二南鉴湖升旗
                    auto& firstColumn = scheduleTable[1][0];// 对应表格三行一列，周一南鉴湖降旗
                    for (auto member : firstColumn) {
                        if (member == person) {
                            return true;
                        }
                    }
                    return false;
                }
                return true;
            }
            case AllHandoverRule: // 全周（周二至周五）南鉴湖升旗采用交接规则
            {
                int currentCol = slot % 2;//值为0~1,检查当前是否为升旗时间段
                if (currentCol == 0 && !location && slot)
                // currentCol == 0:当前为升旗任务  !location == 1:当前为南鉴湖任务  slot != 0:当前不是周一升旗任务
                // 所以，能进入if语句内的条件是：周二到周五的南鉴湖升旗任务
                {
                    auto& prevColumn = scheduleTable[slot-1][location];//前一天南鉴湖降旗情况
                    // 查看该队员是否安排在前一天南鉴湖降旗任务中
                    for (auto member : prevColumn) {
                        if (member == person) {
                            return true;
                        }
                    }
                    return false;
                }
                return true;
            }
            case NoRule: // 不采用交接规则
                return true;
        }
        return true;
    }
};

inline bool SchedulingManager::getUseTotalTimesRule() const
{
    return useTotalTimesRule;
}

inline void SchedulingManager::setUseTotalTimesRule(bool newUseTotalTimesRule)
{
    useTotalTimesRule = newUseTotalTimesRule;
}

inline const Flag_group &SchedulingManager::getFlagGroup() const
{
    return flagGroup;
}

inline std::vector<Person *> SchedulingManager::getAvailableMembers() const
{
    return availableMembers;
}

inline void SchedulingManager::setAvailableMembers(const std::vector<Person *> &newAvailableMembers)
{
    availableMembers = newAvailableMembers;
}

inline std::vector<std::vector<std::vector<Person *> > > SchedulingManager::getScheduleTable() const
{
    return scheduleTable;
}

inline void SchedulingManager::setScheduleTable(const std::vector<std::vector<std::vector<Person *> > > &newScheduleTable)
{
    scheduleTable = newScheduleTable;
}

inline SchedulingManager::HandoverRule SchedulingManager::getHandoverRule() const
{
    return handoverRule;
}

inline void SchedulingManager::setHandoverRule(SchedulingManager::HandoverRule newHandoverRule)
{
    handoverRule = newHandoverRule;
}
