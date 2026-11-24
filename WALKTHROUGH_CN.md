# 本地运行与讲解验收

先用 VS Code 打开这个项目文件夹，在 Terminal 中运行：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

预期 8 项测试通过。然后运行：

```bash
python scripts/reverse_dcf.py
```

打开工作簿，依次看 Historicals → Model_Base → Summary。

## 三个亲手实验

1. 找出 2025 年收入的三部分：器械及配件 6,018.9、系统 2,473.7、服务 1,572.1（百万美元）。相加必须等于 10,064.7。再说出预测时各部分分别由哪些业务变量驱动。
2. 在 Excel 把 Model_Base 的 B18（WACC）从 9% 改为 10%。Summary 的 Base 价值应该下降；改回 9%。B19（终值增长）不能大于等于 WACC。
3. 把 Model_Base 的 B21 从 0 改为 1。解释为什么加回 SBC 会提高模型价值，以及为什么这不等于股东没有付出稀释成本。改回 0。

## 时间口径

这是 2026 年末的前瞻估值，不是 11 月报价日的即期公允价值。Q4 现金桥采用全年 FCF 的 25% 近似。保留的 407.24 美元是原项目给定的历史基准输入，尚未独立验证，不能称为当前报价。

## Reverse DCF

它反问“需要怎样的增长假设才能让模型价值等于给定价格”。它不是预测实际一定会实现这样的增长。Summary 中反向增长两项是脚本生成时的结果；改任何假设后请重跑数值求解，或在 Excel 用 Goal Seek 调整 Model_Base!B20，使 B65 等于基准价格。

多家公司可比估值还没有完整核对，不要把 Comps 模板写成已完成的研究结论。
