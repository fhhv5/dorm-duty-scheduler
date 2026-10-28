# 排班小程序

Windows 离线排班工具。项目发布与维护：wufong（GitHub [@fhhv5](https://github.com/fhhv5)）；开发协助：OpenAI Codex。

从 [发布页](https://github.com/fhhv5/dorm-duty-scheduler/releases/latest) 下载单文件 `scheduling-app.exe`。双击后分别导出名单、班次和冲突模板，填写后导入程序并生成 Excel。使用说明和用户协议可在程序内查看；开源许可可在“关于与感谢”中查看。可选择导出电话、宿舍、签名栏、班级、专业和辅导员；名单模板包含这些人员信息列，均可留空。详细格式、排班规则及使用方法见 [使用说明](使用说明.txt)，请阅读[用户协议](用户协议.txt)。示例为虚构数据，正式使用前请替换。

导入 Excel 时优先读取名为“数据”的工作表；其他 Excel 则读取当前活动表。CSV 使用 UTF-8 编码。程序尽量缩小各成员排班次数的最大差距，且要求班次人数满足、时间不重叠、避开指定冲突。结果请人工核对。

## 源代码与构建

Python 3.12，依赖 `openpyxl==3.1.5`、`ortools==9.15.6755`；界面使用 Python 自带 tkinter。Windows 构建使用 PyInstaller 6.22.3。在项目根目录依次运行 `python 源代码/准备内置资料.py` 和 `python -m PyInstaller --noconfirm 源代码/排班程序.spec --distpath . --workpath 构建临时`。准备资料时会打包当前模板，发布前务必确认其中没有个人数据。模板生成脚本使用 OpenAI 工作区 artifact 工具，运行程序不需要该工具。

## 致谢与许可

项目代码独立编写；参考 [Google OR-Tools 排班示例](https://github.com/google/or-tools/blob/stable/examples/python/shift_scheduling_sat.py) 和 [shift-scheduling-optimizer](https://github.com/Shweird/shift-scheduling-optimizer) 的约束排班思路，并未复制后者代码。使用 OR-Tools、openpyxl 等开源依赖；对应许可和声明见 `第三方许可`。本项目原创代码采用 MIT 许可，第三方组件各自许可继续适用。
