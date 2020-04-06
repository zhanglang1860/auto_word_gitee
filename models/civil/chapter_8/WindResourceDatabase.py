import math
import pandas as pd
from connect_sql import connect_sql_pandas
import numpy as np


# from docxtpl import DocxTemplate
# from RoundUp import round_dict_numbers


class WindResourceDatabase:

    def __init__(self):
        # ===========selecting parameters=============
        self.FortificationIntensity, self.BasicType, self.UltimateLoad = 0, "", 0
        # ===========basic parameters==============
        self.data_wind_resource, self.DataWindResource = pd.DataFrame(), pd.DataFrame()
        self.basic_earthwork_ratio, self.basic_stone_ratio, self.turbine_numbers = 0, 0, 0
        self.dict_wind_resource, self.dict_wind_resource_02 = {}, {}
        # ===========Calculated parameters==============
        self.earth_excavation_wind_resource, self.stone_excavation_wind_resource = 0, 0
        self.earth_work_back_fill_wind_resource, self.earth_excavation_wind_resource_numbers = 0, 0
        self.stone_excavation_wind_resource_numbers, self.earth_work_back_fill_wind_resource_numbers = 0, 0
        self.c40_wind_resource_numbers, self.c15_wind_resource_numbers, self.c80_wind_resource_numbers = 0, 0, 0
        self.c80_wind_resource_numbers, self.reinforcement_wind_resource_numbers = 0, 0

    def extraction_data_wind_resource(self, fortification_intensity, basic_type, ultimate_load):
        self.FortificationIntensity = fortification_intensity
        self.BasicType = basic_type
        self.UltimateLoad = ultimate_load
        # col_name = ["FortificationIntensity", "BearingCapacity", "BasicType", "UltimateLoad", "FloorRadiusR", "R1",
        #             "R2", "H1", "H2", "H3", "PileDiameter", "Number", "Length", "SinglePileLength", "Area", "Volume",
        #             "Cushion", "M48PreStressedAnchor", "C80SecondaryGrouting"]
        #
        # self.DataWindResource = pd.read_excel(
        #     r"C:\Users\Administrator\PycharmProjects\Odoo_addons_NB\autocrword\models\chapter_8\chapter8database.xlsx",
        #     header=1, sheet_name="风机基础数据", usecols=col_name)

        sql = "SELECT * FROM auto_word_civil_windbase"
        self.DataWindResource = connect_sql_pandas(sql)
        self.data_wind_resource = \
            self.DataWindResource.loc[
                self.DataWindResource["FortificationIntensity"] == self.FortificationIntensity].loc[
                self.DataWindResource["BasicType"] == self.BasicType].loc[
                self.DataWindResource["UltimateLoad"] == self.UltimateLoad]
        return self.data_wind_resource

    def excavation_cal_wind_resource(self, data_wind_resource, basic_earthwork_ratio, basic_stone_ratio, turbine_num):
        self.data_wind_resource = data_wind_resource
        self.basic_earthwork_ratio = basic_earthwork_ratio
        self.basic_stone_ratio = basic_stone_ratio
        self.turbine_numbers = turbine_num

        self.earth_excavation_wind_resource = \
            math.pi * (self.data_wind_resource["FloorRadiusR"] + 1.3) ** 2 * \
            (self.data_wind_resource["H1"] + self.data_wind_resource["H2"] + self.data_wind_resource["H3"] + 0.15) \
            * self.basic_earthwork_ratio

        self.stone_excavation_wind_resource = \
            math.pi * (self.data_wind_resource["FloorRadiusR"] + 1.3) ** 2 * \
            (self.data_wind_resource["H1"] + self.data_wind_resource["H2"] + self.data_wind_resource["H3"] + 0.15) \
            * self.basic_stone_ratio

        self.earth_work_back_fill_wind_resource = \
            self.earth_excavation_wind_resource + self.stone_excavation_wind_resource - \
            self.data_wind_resource["Volume"] - self.data_wind_resource["Cushion"]

        self.stone_excavation_wind_resource_numbers = self.stone_excavation_wind_resource * int(self.turbine_numbers)
        self.earth_excavation_wind_resource_numbers = self.earth_excavation_wind_resource * self.turbine_numbers
        self.earth_work_back_fill_wind_resource_numbers = self.earth_work_back_fill_wind_resource * self.turbine_numbers

        print(self.data_wind_resource.at[self.data_wind_resource.index[0], "Volume"])
        print(self.turbine_numbers)
        self.c40_wind_resource_numbers = \
            self.data_wind_resource.at[self.data_wind_resource.index[0], "Volume"] * self.turbine_numbers
        self.c15_wind_resource_numbers = \
            self.data_wind_resource.at[self.data_wind_resource.index[0], "Cushion"] * self.turbine_numbers
        self.c80_wind_resource_numbers = \
            self.data_wind_resource.at[self.data_wind_resource.index[0], "C80SecondaryGrouting"] * \
            self.turbine_numbers
        self.data_wind_resource["EarthExcavation_WindResource"] = self.earth_excavation_wind_resource
        self.data_wind_resource["StoneExcavation_WindResource"] = self.stone_excavation_wind_resource
        self.data_wind_resource["EarthWorkBackFill_WindResource"] = self.earth_work_back_fill_wind_resource
        self.data_wind_resource["Reinforcement"] = self.data_wind_resource["Volume"] * 0.1
        self.reinforcement_wind_resource_numbers = \
            self.data_wind_resource.at[self.data_wind_resource.index[0], "Reinforcement"] * self.turbine_numbers

        return self.data_wind_resource

    def generate_dict_wind_resource(self, data, turbine_num):
        self.data_wind_resource = data
        self.turbine_numbers = turbine_num
        self.dict_wind_resource = {
            "turbine_numbers": int(self.turbine_numbers),
            "土方开挖_风机": self.data_wind_resource.at[self.data_wind_resource.index[0], "EarthExcavation_WindResource"],
            "石方开挖_风机": self.data_wind_resource.at[self.data_wind_resource.index[0], "StoneExcavation_WindResource"],
            "土石方回填_风机": self.data_wind_resource.at[self.data_wind_resource.index[0], "EarthWorkBackFill_WindResource"],
            "C40混凝土_风机": self.data_wind_resource.at[self.data_wind_resource.index[0], "Volume"],
            "C15混凝土_风机": self.data_wind_resource.at[self.data_wind_resource.index[0], "Cushion"],
            "钢筋_风机": self.data_wind_resource.at[self.data_wind_resource.index[0], "Reinforcement"],
            "基础防水_风机": 1,
            "沉降观测_风机": 4,
            "预制桩长_风机": self.data_wind_resource.at[self.data_wind_resource.index[0], "SinglePileLength"],
            "M48预应力锚栓_风机": self.data_wind_resource.at[self.data_wind_resource.index[0], "M48PreStressedAnchor"],
            "C80二次灌浆_风机": self.data_wind_resource.at[self.data_wind_resource.index[0], "C80SecondaryGrouting"],
            "单台风机桩根数_风机": self.data_wind_resource.at[self.data_wind_resource.index[0], "Number"],
        }
        self.dict_wind_resource_02 = {
            "基础形式": self.data_wind_resource.at[self.data_wind_resource.index[0], "BasicType"],
            "基础底面圆直径": self.data_wind_resource.at[self.data_wind_resource.index[0], "FloorRadiusR"]*2,
            "基础底板外缘高度": self.data_wind_resource.at[self.data_wind_resource.index[0], "H1"],
            "台柱圆直径": self.data_wind_resource.at[self.data_wind_resource.index[0], "R2"]*2,
            "基础底板圆台高度": self.data_wind_resource.at[self.data_wind_resource.index[0], "H2"],
            "台柱高度": self.data_wind_resource.at[self.data_wind_resource.index[0], "H3"],
            "抗震烈度": self.data_wind_resource.at[self.data_wind_resource.index[0], "FortificationIntensity"],
        }
        return self.dict_wind_resource,self.dict_wind_resource_02

# project01 = WindResourceDatabase()
# data = project01.extraction_data_wind_resource(basic_type="扩展基础", ultimate_load=70000, fortification_intensity=7)
# turbine_numbers = 15
# data_cal = project01.excavation_cal_wind_resource(data,0.8, 0.2, turbine_numbers)
# dict_wind_resource = project01.generate_dict_wind_resource(data_cal, turbine_numbers)
# Dict = round_dict_numbers(dict_wind_resource, dict_wind_resource["numbers_tur"])
#
# docx_box = ["cr8", "result_chapter8"]
# save_path = r"C:\Users\Administrator\PycharmProjects\Odoo_addons_NB\autocrword\models\chapter_8"
# readpath = os.path.join(save_path, "%s.docx") % docx_box[0]
# savepath = os.path.join(save_path, "%s.docx") % docx_box[1]
# tpl = DocxTemplate(readpath)
# tpl.render(Dict)
# tpl.save(savepath)
