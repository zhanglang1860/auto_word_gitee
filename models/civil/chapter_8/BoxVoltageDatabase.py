import pandas as pd
from connect_sql import connect_sql_pandas


# from RoundUp import round_dict_numbers
# from docxtpl import DocxTemplate

class BoxVoltageDatabase:
    def __init__(self):
        # ===========selecting parameters=============
        self.TurbineCapacity = 0
        # ===========basic parameters==============
        self.data_box_voltage, self.DataBoxVoltage = pd.DataFrame(), pd.DataFrame()
        self.basic_earthwork_ratio, self.basic_stone_ratio, self.turbine_numbers = 0, 0, 0
        self.dict_box_voltage = {}
        # ===========Calculated parameters==============
        self.earth_excavation_box_voltage, self.stone_excavation_box_voltage = 0, 0
        self.earthwork_back_fill_box_voltage, self.earth_excavation_box_voltage_numbers = 0, 0
        self.stone_excavation_box_voltage_numbers, self.earthwork_back_fill_box_voltage_numbers = 0, 0

        self.c35_box_voltage_numbers, self.c15_box_voltage_numbers, self.reinforcement_box_voltage_numbers = 0, 0, 0

    def extraction_data_box_voltage(self, turbine_capacity):
        self.TurbineCapacity = str(int(turbine_capacity))
        # col_name = ["TurbineCapacity", "ConvertStation", "Long", "Width", "High", "WallThickness", "HighPressure",
        #             "C35ConcreteTop", "C15Cushion", "MU10Brick", "Reinforcement", "Area"]
        # self.DataBoxVoltage = pd.read_excel(
        #     r"C:\Users\Administrator\PycharmProjects\Odoo_addons_NB\autocrword\models\chapter_8\chapter8database.xlsx",
        #     header=2, sheet_name="箱变基础数据", usecols=col_name)

        sql = "SELECT * FROM auto_word_civil_convertbase"
        self.DataBoxVoltage = connect_sql_pandas(sql)

        self.data_box_voltage = self.DataBoxVoltage.loc[self.DataBoxVoltage["TurbineCapacity"] == self.TurbineCapacity]
        return self.data_box_voltage

    def excavation_cal_box_voltage(self, data_box_voltage, basic_earthwork_ratio, basic_stone_ratio, turbine_num):
        self.data_box_voltage = data_box_voltage
        self.basic_earthwork_ratio = basic_earthwork_ratio
        self.basic_stone_ratio = basic_stone_ratio
        self.turbine_numbers = turbine_num

        self.earth_excavation_box_voltage = \
            (self.data_box_voltage["Long"] + 0.5 * 2) * (self.data_box_voltage["Width"] + 0.5 * 2) * \
            (self.data_box_voltage["High"] - 0.2) * self.basic_earthwork_ratio
        self.stone_excavation_box_voltage = \
            (self.data_box_voltage["Long"] + 0.5 * 2) * (self.data_box_voltage["Width"] + 0.5 * 2) * \
            (self.data_box_voltage["High"] - 0.2) * self.basic_stone_ratio
        self.earthwork_back_fill_box_voltage = \
            self.earth_excavation_box_voltage + self.stone_excavation_box_voltage - self.data_box_voltage["Long"] * \
            self.data_box_voltage["Width"] * (self.data_box_voltage["High"] - 0.2)

        self.data_box_voltage = self.data_box_voltage.copy()

        self.earth_excavation_box_voltage_numbers = self.earth_excavation_box_voltage * self.turbine_numbers
        self.stone_excavation_box_voltage_numbers = self.stone_excavation_box_voltage * self.turbine_numbers
        self.earthwork_back_fill_box_voltage_numbers = self.earthwork_back_fill_box_voltage * self.turbine_numbers

        self.c35_box_voltage_numbers = \
            self.data_box_voltage.at[self.data_box_voltage.index[0], "C35ConcreteTop"] * self.turbine_numbers
        self.c15_box_voltage_numbers = \
            self.data_box_voltage.at[self.data_box_voltage.index[0], "C15Cushion"] * self.turbine_numbers

        self.reinforcement_box_voltage_numbers = \
            self.data_box_voltage.at[self.data_box_voltage.index[0], "Reinforcement"] * self.turbine_numbers

        self.data_box_voltage["EarthExcavation_BoxVoltage"] = self.earth_excavation_box_voltage
        self.data_box_voltage["StoneExcavation_BoxVoltage"] = self.stone_excavation_box_voltage
        self.data_box_voltage["EarthWorkBackFill_BoxVoltage"] = self.earthwork_back_fill_box_voltage

        return self.data_box_voltage

    def generate_dict_box_voltage(self, data_box_voltage, turbine_num):
        self.data_box_voltage = data_box_voltage
        self.turbine_numbers = turbine_num
        self.dict_box_voltage = {
            "numbers_box_voltage": int(self.turbine_numbers),
            "土方开挖_箱变": self.data_box_voltage.at[self.data_box_voltage.index[0], "EarthExcavation_BoxVoltage"],
            "石方开挖_箱变": self.data_box_voltage.at[self.data_box_voltage.index[0], "StoneExcavation_BoxVoltage"],
            "土石方回填_箱变": self.data_box_voltage.at[self.data_box_voltage.index[0], "EarthWorkBackFill_BoxVoltage"],
            "C35混凝土_箱变": self.data_box_voltage.at[self.data_box_voltage.index[0], "C35ConcreteTop"],
            "C15混凝土_箱变": self.data_box_voltage.at[self.data_box_voltage.index[0], "C15Cushion"],
            "MU10砖_箱变": self.data_box_voltage.at[self.data_box_voltage.index[0], "MU10Brick"],
            "钢筋_箱变": self.data_box_voltage.at[self.data_box_voltage.index[0], "Reinforcement"],
        }
        return self.dict_box_voltage

# turbine_numbers = 15
# turbine_capacity=3
# project02 = BoxVoltageDatabase()
# data = project02.extraction_data_box_voltage(turbine_capacity)
# data_cal = project02.excavation_cal_box_voltage(0.8, 0.2, turbine_numbers)
#
# dict_box_voltage = project02.generate_dict_box_voltage(data_cal, numbers_list)
# Dict = round_dict_numbers(dict_box_voltage, dict_box_voltage["numbers_box_voltage"])
#
# docx_box = ["cr8", "result_chapter8"]
# save_path = r"C:\Users\Administrator\PycharmProjects\Odoo_addons_NB\autocrword\models\chapter_8"
# readpath = os.path.join(save_path, "%s.docx") % docx_box[0]
# savepath = os.path.join(save_path, "%s.docx") % docx_box[1]
# tpl = DocxTemplate(readpath)
# tpl.render(Dict)
# tpl.save(savepath)
