# flightsheet_process
基于python的每日进出港数据生成程序，用于排序筛选所需航班及根据停机位获取机位备注

该项目基于某机场每日统计航班数据的需求：

- 只选择合约航班
- 需要根据航班号获取前两位航司代码
- 按照时间排序
- 9C为开头的航班需要获取配载备注的数字部分以获得该航班人数
- 不同区域机位需要进行机位备注：T3近机位为1 卫星厅为5 远机位为2 国际航班及地区航班为3

该程序基于python，使用pandas库对excel数据进行获取与处理，使用者仅需在弹出窗口中选择相应的数据（存在格式要求），程序自动处理并弹出保存窗口，保存即可。

**该项目仍然处于测试阶段**
