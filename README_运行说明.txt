C题四问代码（openpyxl 本地版）
================================

一、改动说明
------------
本版本已彻底删除 artifact_tool 依赖。
Excel 的读取、模板填充与保存全部改为 openpyxl。
数学模型、岭回归、MILP、储能执行逻辑保持不变。

二、安装依赖
------------
在 PowerShell / CMD 中进入本文件夹后运行：

    python -m pip install -r requirements.txt

或者：

    python -m pip install numpy scipy openpyxl

注意：SciPy 版本需要支持 scipy.optimize.milp，建议 SciPy >= 1.11。

三、文件放置
------------
请把以下原始数据和模板放在与代码相同的文件夹中：

    附件1(3).xlsx
    附件2(3).xlsx
    附件3(3).xlsx
    附件4(3).xlsx
    result1(9).xlsx
    result2(6).xlsx
    result3(6).xlsx
    result4-2(6).xlsx
    result4-3(6).xlsx

四、分别运行四问
----------------
问题1：
    python q1_solution.py

问题2：
    python q2_solution.py

问题3：
    python q3_solution.py

问题4：
    python q4_solution.py

五、输出文件
------------
问题1 -> result1.xlsx
问题2 -> result2.xlsx
问题3 -> result3.xlsx
问题4 -> result4-2.xlsx、result4-3.xlsx

六、依赖测试
------------
    python -c "import numpy, scipy, openpyxl; from scipy.optimize import milp; print('依赖正常')"

七、常见报错
------------
1. ModuleNotFoundError: No module named 'openpyxl'
   运行：python -m pip install openpyxl

2. ImportError: cannot import name 'milp' from scipy.optimize
   运行：python -m pip install -U scipy

3. FileNotFoundError
   检查附件和 result 模板文件名是否与上面完全一致，并且与代码位于同一目录。
