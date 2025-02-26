#include "Flag_group.h"

// 添加队员到指定组
void Flag_group::addPersonToGroup(const Person &person, int groupNumber)
{
    // 参数：Person类：待添加的队员信息。 int groupNumber：队员对应的组别。
    // 根据组别将新队员person加入到对应的组中
    if (groupNumber >= 1 && groupNumber <= 4) // 判断组号是否合理：1~4对应一至四组
    {
        // 因为group索引最小为0，与输入组号存在一位的差距，需要减一处理
        group[groupNumber - 1].push_back(person);// push_back,添加新队员至对应组
    }
    // 测试代码
    // else
    // {
    //     std::cerr << "非法组号，所属组名应为1~4" << endl;
    // }
}

// 从指定组中删除指定的队员
void Flag_group::removePersonFromGroup(const Person &person, int groupNumber)
{
    // 参数：Person类：待删除的队员信息。 int groupNumber：队员对应的组别。
    // 将根据参数的groupNumber在对应的组中查找是否存在参数中的person，查找到后，删除队员。
    if (groupNumber >= 1 && groupNumber <= 4)
    {
        // 因为group索引最小为0，与输入组号存在一位的差距，需要减一处理
        vector<Person> &currentGroup = group[groupNumber - 1];
        // 从对应组中依次查找是否存在要删除的队员
        for (auto it = currentGroup.begin(); it != currentGroup.end(); ++it)
        {
            if (it->getName() == person.getName() && it->getGender() == person.getGender() && it->getClassname() == person.getClassname())
            {
                // 不能通过组名判断person是否是要删除的那个队员，因为在onGroupComboBoxChanged函数中已经更新了person的组别
                currentGroup.erase(it);
                return; // 查找到队员后，终止查找
            }
        }
        // 测试代码
        // std::cerr << "removePersonFromGroup()未找到" << person.getName() << endl;
    }
    // 测试代码
    // else
    // {
    //     std::cerr << "非法组号，所属组名应为1~4" << endl;
    // }
}

// 修改指定组中指定姓名的队员信息
void Flag_group::modifyPersonInGroup(const Person& oldPerson, const Person& newPerson, int groupNumber)
{
    // 参数：Person类：oldPerson:待修改的队员，newPerson：用于替换原队员信息的新信息。 int groupNumber：队员对应的组别。
    // 将根据参数的groupNumber在对应的组中查找是否存在参数中的person，查找到后，用newPerson中数据替换oldPerson的数据，完成修改
    // 修改组别以外的队员信息，如果是组员修改组别信息将不从此函数进行
    if (groupNumber >= 1 && groupNumber <= 4) {
        // 因为group索引最小为0，与输入组号存在一位的差距，需要减一处理
        vector<Person>& currentGroup = group[groupNumber - 1];
        // 从对应组中依次查找是否存在待修改的队员
        for (auto& person : currentGroup) {
            if (person == oldPerson) {
                person = newPerson;
                return;
            }
        }
        // 测试代码
        // std::cerr << "modifyPersonInGroup()未找到" << oldPerson.getName() << endl;
    }
    // 测试代码
    // else
    // {
    //     std::cerr << "非法组号，所属组名应为1~4" << endl;
    // }

}

// 在指定组中查找指定姓名的队员
Person* Flag_group::findPersonInGroup(const Person &person, int groupNumber)
{
    // 参数：Person类：待查找的队员信息。 int groupNumber：队员对应的组别。
    // 将根据参数的groupNumber在对应的组中查找是否存在参数中的person，查找到后，返回该队员
    if (groupNumber >= 1 && groupNumber <= 4)
    {
        vector<Person> &currentGroup = group[groupNumber - 1];
        for (auto &tempPerson : currentGroup)
        {
            if (tempPerson.getName() == person.getName() && tempPerson.getGender() == person.getGender() && tempPerson.getClassname() == person.getClassname())
            {
                // 不能通过组名判断person是否是要查找的那个队员，因为在onGroupComboBoxChanged函数中已经更新了person的组别(即切换队员组别功能）
                return &tempPerson;
            }
        }
        // 测试代码
        // std::cerr << "findPersonInGroup()未找到" << person.getName() << endl;
    }
    // 测试代码
    // else
    // {
    //     std::cerr << "非法组号，所属组名应为1~4" << endl;
    // }
    return nullptr;
}

// 在全队查找指定姓名的队员
Person* Flag_group::findPerson(const Person &person) {
    // 参数：Person类：待查找的队员信息。
    for (int i = 1; i <= 4; ++i) {
        Person* found = findPersonInGroup(person, i);
        if (found) {
            return found;
        }
    }
    // 测试代码
    // std::cerr << "findPerson()未找到" << person.getName() << endl;
    return nullptr;
}

// 获取指定组的所有队员
// 非常量版本，允许修改返回的向量
vector<Person>& Flag_group::getGroupMembers(int groupNumber)
{
    // 参数：int groupNumber：待遍历的组别。
    if (groupNumber >= 1 && groupNumber <= 4)
    {
        return group[groupNumber - 1];
    }
    else
    {
        static std::vector<Person> emptyGroup;
        // 测试代码
        // std::cerr << "非法组号，所属组名应为1~4" << std::endl;
        return emptyGroup;
    }
}
// 获取指定组的所有队员
// 常量版本，用于只读访问
const vector<Person>& Flag_group::getGroupMembers(int groupNumber) const
{
    // 参数：int groupNumber：待遍历的组别。
    if (groupNumber >= 1 && groupNumber <= 4)
    {
        return group[groupNumber - 1];
    }
    else
    {
        static vector<Person> emptyGroup;
        // 测试代码
        // std::cerr << "非法组号，所属组名应为1~4" << endl;
        return emptyGroup;
    }
}
