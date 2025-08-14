# config.py - 配置文件
# ⚠️ 请根据你的实际需求修改以下配置

class Config:
    # 📂 文件路径配置
    SOURCE_DIR = "source_reports"  # 原始报告文件夹
    TEMPLATE_DIR = "templates"  # 模板文件夹
    OUTPUT_DIR = "output"  # 输出文件夹
    TEMPLATE_FILE = "template_report.xlsx"  # 模板文件名

    # 📊 数据源配置 - ⚠️ 根据你的Excel表结构修改
    SOURCE_SHEET_NAME = "Data"  # 原始数据表名
    TEMPLATE_SHEET_NAMES = ["HTRB 100%", "AC"]  # 模板中的表名列表

    # 📍 原报告数据位置配置 - ⚠️ 根据你的原报告格式修改
    SOURCE_DATA_POSITIONS = {
        "item_name_row": 15,  # 测试项目名称行号 (Item Name行)
        "bias1_row": 16,  # Bias1行号
        "bias2_row": 17,  # Bias2行号
        "bias3_row": 18,  # Bias3行号
        "min_limit_row": 19,  # Min Limit行号
        "max_limit_row": 20,  # Max Limit行号
        "data_start_row": 28,  # 测试数据开始行号 (P1开始的行)
        "data_start_col": 2,  # 数据开始列号 (通常是A列，包含P1,P2...)
        "test_items_start_col": 2,  # 测试项目开始列号 (通常是B列)
        "sample_id_col": 0  # 样品ID列号 (P1, P2, P3...所在列)
    }

    # 🔍 数据识别配置 - ⚠️ 根据你的数据格式修改
    DATA_RECOGNITION = {
        "sample_prefix": "P",  # 样品编号前缀 (如P1, P2中的P)
        "skip_empty_rows": True,  # 是否跳过空行
        "auto_detect_data_end": True,  # 是否自动检测数据结束
        "max_data_rows": 100,  # 最大数据行数 (防止读取过多无效数据)
        # 🆕 新增：特殊值识别
        "over_value_patterns": ["OVER", "Over", "over"],  # "Over"值的识别模式
        "treat_over_as_abnormal": True  # 是否将"Over"值视为异常
    }

    # 🎯 测试项目映射 - ⚠️ 根据你的测试项目修改
    TEST_ITEMS_MAPPING = {
        "HVISG": ["5 ISGS"],  # 高压漏电流
        "VGS(th)": ["7 VTH"],  # 阈值电压
        "BVDSS": ["8 BVDSS"],  # 击穿电压
        "HVIDSS": ["9 IDSS", "9 HVIDSS"],  # 高压漏电流
        "RDS(ON)": ["10 RDON"]  # 导通电阻
    }

    # 📋 数据分组配置 - ⚠️ 根据你的数据分组需求修改
    DATA_GROUPS = {
        "group1": {
            "range": (1, 22),  # P1-P22
            "target_sheet": 1,  # 写入模板的第一个表
            "description": "第一批次数据"
        },
        "group2": {
            "range": (23, 44),  # P23-P44
            "target_sheet": 2,  # 写入模板的第二个表
            "description": "第二批次数据"
        }
    }

    # 🎨 格式配置
    HIGHLIGHT_COLOR = "FFFF00"  # 黄色高亮颜色
    # 🆕 新增：Over值的特殊高亮颜色
    OVER_VALUE_HIGHLIGHT_COLOR = "FF6B6B"  # 红色高亮，用于Over值

    # 📍 模板位置配置 - ⚠️ 根据你的模板布局修改
    TEMPLATE_POSITIONS = {
        "test_items_row": 8,  # 测试项目行号
        "test_conditions_row": 10,  # 测试条件起始行
        "test_conditions_max_rows": 2,  # ⚠️ 新增：测试条件最大行数，如果有多个条件会占用多行
        "min_limit_row": 12,  # 规格下限行号
        "max_limit_row": 13,  # 规格上限行号
        "data_start_row": 18,  # 数据开始行号
        "test_items_start_col": 2,  # 测试项目开始列号(B列)
        "sample_id_col": 1  # 样品序号列(A列)
    }

    # 🔧 数据处理配置 - ⚠️ 根据需要修改
    DATA_PROCESSING = {
        "convert_to_numeric": True,  # ⚠️ 新增：是否转换为数值形式
        "numeric_precision": 6,  # ⚠️ 新增：数值精度
        "conditions_multiline": True,  # ⚠️ 新增：测试条件是否分行显示
        "combine_conditions_separator": "; ",
        "combine_values_separator": "; ",
        "empty_value_placeholder": "",
        "invalid_value_placeholder": "N/A",
        # 🆕 新增：Over值处理配置
        "preserve_over_values": True,  # 保持Over值原样输出
        "over_value_display": "Over"  # Over值的显示格式
    }

    # 📊 数值处理配置 - ⚠️ 增强数值转换功能
    VALUE_PROCESSING = {
        "unit_patterns": {
            "voltage": ["V", "mV"],
            "current": ["A", "mA", "uA", "nA"],
            "resistance": ["R", "mR", "ohm", "Ω"]
        },
        # ⚠️ 修复：注释掉单位转换，因为数据和限值单位一致
        # "unit_conversions": {
        #     "mV": 0.001,
        #     "uA": 0.000001,
        #     "nA": 0.000000001,
        #     "mR": 0.001
        # },
        "decimal_places": 6,  # ⚠️ 修改：增加精度
        "scientific_notation_threshold": 0.000001,
        "remove_units_for_comparison": True,
        "force_numeric_output": True  # ⚠️ 新增：强制数值输出
    }

    # 🆕 新增：异常统计配置
    ABNORMAL_STATISTICS = {
        "enable_counting": True,  # 是否启用异常统计
        "count_per_row": True,  # 同一行多个异常只计算一次
        "include_over_values": True,  # 将Over值计入异常统计
        "write_to_template": True,  # 是否将统计结果写入模板

        # 异常统计写入位置配置 - ⚠️ 根据你的模板调整
        "positions": {
            "sheet_1": {  # 第一个工作表
                "row": 40,  # 写入行号
                "col": 2,  # 写入列号
                "write_as_number": True,  # 🆕 写入数值而不是文本
                "format": "{count}"  # 显示格式
            },
            "sheet_2": {  # 第二个工作表
                "row": 40,  # 写入行号
                "col": 2,  # 写入列号
                "write_as_number": True,  # 🆕 写入数值而不是文本
                "format": "{count}"  # 显示格式
            }
        },

        # 异常统计的详细配置
        "detailed_statistics": {
            "enable": False,  # 是否启用详细统计
            "by_test_item": False,  # 按测试项统计
            "by_sample": False,  # 按样品统计
            "export_to_separate_sheet": False  # 导出到单独的工作表
        }
    }

    # 🚨 错误处理配置
    ERROR_HANDLING = {
        "continue_on_error": True,
        "log_detailed_errors": True,
        "create_error_report": True,
        "backup_original": False
    }

    # 🆕 新增：日志配置
    LOGGING = {
        "level": "DEBUG",  # 日志级别
        "log_abnormal_details": True,  # 记录异常详情
        "log_over_value_detection": True,  # 记录Over值检测
        "log_statistics": True  # 记录统计信息
    }

    # 🆕 新增：验证配置
    VALIDATION = {
        "check_template_structure": True,  # 检查模板结构
        "validate_data_ranges": True,  # 验证数据范围
        "warn_missing_limits": True,  # 警告缺失的限值
        "strict_mode": False  # 严格模式（遇到错误停止处理）
    }
