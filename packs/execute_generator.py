import xlrd
import xlwt
from packs.rule import Rule


class ExecuteGenerator:
    def __init__(self, source_file, template_file, rule_file, output_folder):
        self.source_file = source_file
        self.template_file = template_file
        self.rule_file = rule_file
        self.output_folder = output_folder

    def parse_source_data_excel(self, rule_dict):
        """
        解析源数据表，返回以项目名称为 key、试验号为 value 的字典。
        """
        result_dict = {}
        wb = xlrd.open_workbook(self.source_file)
        # 取出文件的第一张表
        sheet = wb.sheet_by_index(0)
        # 取出项目名称作为key
        keys = sheet.col_values(3)
        # 取出试验号作为value
        values = sheet.col_values(4)
        # 将所有数据写入字典
        [
            result_dict.update({keys[i]: [values[i]]})
            if keys[i] not in result_dict
            else result_dict[keys[i]].append(values[i])
            for i in range(1, len(keys))
        ]

        for key in list(result_dict.keys()):
            value_list = result_dict[key]
            # 数据量大时需要分表，具体分表条件看是否有3个NC
            limit = (
                267
                if rule_dict.get(key, {}).get("three_NC", False)
                else 270
                if key in rule_dict
                else 270
            )
            if len(value_list) > limit:
                print(f"类型：{key} 对应的 value 长度 > {limit}")
                # 将列表分成270个一份，组合成数组
                sub_arrays = [
                    value_list[x: x + limit] for x in range(0, len(value_list), limit)
                ]
                # 使用字典推导式更新原字典
                result_dict = {
                    **result_dict,
                    **{f"{key}_{i}": sub_arrays[i] for i in range(1, len(sub_arrays))},
                }
                # 只保留第一个分表
                result_dict[key] = sub_arrays[0]

        return result_dict

    def parse_rule_excel(self):
        """
        解析规则表，返回规则字典。
        """
        wb = xlrd.open_workbook(self.rule_file)
        # 取出文件的第一张表
        sheet = wb.sheet_by_index(0)
        # 取出项目名称作为key
        rule_keys = sheet.col_values(0)
        # 取出后面的规则放入object作为value
        rule_dict = {
            rule_keys[i]: Rule(
                sheet.cell_value(i, 1),
                sheet.cell_value(i, 2),
                bool(sheet.cell_value(i, 3)),
                sheet.cell_value(i, 4),
                sheet.cell_value(i, 5),
                sheet.cell_value(i, 6),
                sheet.cell_value(i, 7),
            )
            for i in range(1, len(rule_keys))
        }
        return rule_dict

    def core(self, result_dict, rule_dict):
        """
        执行操作表生成任务。
        :param result_dict:结果字典
        :param rule_dict:规则字典
        """
        for key, values in result_dict.items():
            # 加载模板 Excel 文件
            wb = xlrd.open_workbook(self.template_file)
            sheet = wb.sheet_by_index(0)
            # 创建一个新的sheet
            new_wb = xlwt.Workbook()
            new_sheet = new_wb.add_sheet("Sheet1", cell_overwrite_ok=True)
            # 向新工作表中写入数据
            for i in range(sheet.nrows):
                for j in range(sheet.ncols):
                    new_sheet.write(i, j, sheet.cell(i, j).value)

            # 取出对应的规则
            a = key.split("_")[0] if "_" in key else key
            if a in rule_dict:
                rule = rule_dict[a]

            # 插入计算方式
            new_sheet.write(8, 0, rule.formula_mode)

            # 插入方块
            for i, diamond in zip(
                range(40, 47, 2),
                [rule.diamond_1, rule.diamond_2, rule.diamond_3, rule.diamond_4],
            ):
                new_sheet.write(i, 0, diamond)

            # 插入英文代号
            for i, j in [(3, 4), (10, 1), (20, 1), (30, 1)]:
                new_sheet.write(i, j, rule.en)

            # 插入试验号&NC
            three_NC = rule.three_NC
            sub_values = self.insert_values(12, 1, values, three_NC, new_sheet)
            sub_values = self.insert_values(
                22, 1, sub_values, three_NC, new_sheet)
            self.insert_values(32, 1, sub_values, three_NC, new_sheet)
            # 将新工作簿保存为一个新的 Excel 文件
            filename = self.output_folder + "/{}.xls".format(f"{key}操作表")
            print("生成表格：" + filename)
            new_wb.save(filename)

    def insert_values(self, row_cnt, col_cnt, values, three_NC, sheet):
        """
        插入数据
        :param row_cnt:行数
        :param col_cnt:列数
        :param values:数组
        :param three_NC:是否存在3个NC
        :param sheet:表格
        :return 刨除数组中已经插入的数据后的子数组
        """
        now_row_cnt = row_cnt
        now_col_cnt = col_cnt
        for i, value in enumerate(values):
            sheet.write(now_row_cnt, now_col_cnt, value)
            now_row_cnt += 1
            if now_row_cnt > row_cnt + 7:
                now_col_cnt += 1
                now_row_cnt = row_cnt
            if (
                not three_NC
                and now_col_cnt > col_cnt + 10
                and now_row_cnt > row_cnt + 1
            ) or (three_NC and now_col_cnt > col_cnt + 10 and now_row_cnt > row_cnt):
                break
        else:
            return []
        if three_NC:
            sheet.write(row_cnt + 1, col_cnt + 11, "NC")
        return values[i + 1:]

    def execute(self):
        print("========开始执行========")
        rule_dict = self.parse_rule_excel()
        result_dict = self.parse_source_data_excel(rule_dict)
        self.core(result_dict, rule_dict)
        print("========执行完毕========")
