from ElectricalCircuit import ElectricalCircuit
from RoundUp import round_up


class TowerType(ElectricalCircuit):
    """
    铁塔

    """

    def __init__(self, *value_list):
        ElectricalCircuit.__init__(self, *value_list)
        self.tower_type, self.tower_type_high = "", ""
        self.used_numbers, self.tower_weight, self.tower_number_weight = 0, 0, 0
        self.used_numbers_single_J2_24, self.used_numbers_single_J4_24, self.used_numbers_single_FS_18 = 0, 0, 0
        self.used_numbers_double_SJ2_24, self.used_numbers_double_SJ4_24, self.used_numbers_single_Z2_30 = 0, 0, 0
        self.used_numbers_single_ZK_42, self.used_numbers_double_SZ2_30, self.used_numbers_double_SZK_42 = 0, 0, 0
        self.used_numbers_angle_steel = 0

        self.tower_number_area = []
        self.tower_height, self.tower_foot_distance, self.tower_area, self.sum_tower_area = 0, 0, 0, 0
        self.sum_used_numbers, self.sum_tower_number_weight, self.kilometer_tower_number = 0, 0, 0

        self.tower_type_list, self.tower_type_high_list, self.tower_weight_list = [], [], []
        self.tower_height_list, self.tower_foot_distance_list = [], []

    def tower_type_models(self, tower_type, tower_type_high, tower_weight, tower_height, tower_foot_distance):
        self.tower_type = tower_type
        self.tower_type_high = tower_type_high
        self.tower_weight = tower_weight
        self.tower_height = tower_height
        self.tower_foot_distance = tower_foot_distance

        if self.project_chapter6_type == 1:
            if self.tower_type == "单回耐张塔":

                if self.tower_type_high == "J2_24" or self.tower_type_high == "J4_24":
                    self.used_numbers = round_up(self.single_circuit / 3, 0)
                    self.tower_number_weight = self.tower_weight * self.used_numbers
                    self.tower_area = self.used_numbers * self.tower_foot_distance ** 2
                    if self.tower_type_high == "J2_24":
                        self.used_numbers_single_J2_24 = self.used_numbers
                    elif self.tower_type_high == "J4_24":
                        self.used_numbers_single_J4_24 = self.used_numbers
                elif self.tower_type_high == "FS_18":
                    self.used_numbers = round_up(self.line_loop_number / 2, 0) + 1
                    self.tower_number_weight = self.tower_weight * self.used_numbers
                    self.tower_area = self.used_numbers * self.tower_foot_distance ** 2
                    self.used_numbers_single_FS_18 = self.used_numbers

            if self.tower_type == "单回直线塔":
                if self.tower_type_high == "Z2_30" or self.tower_type_high == "ZK_42":
                    self.used_numbers = round_up(self.single_circuit * 1000 / 250 / 2, 0)
                    self.tower_number_weight = self.tower_weight * self.used_numbers
                    self.tower_area = self.used_numbers * self.tower_foot_distance ** 2
                    if self.tower_type_high == "Z2_30":
                        self.used_numbers_single_Z2_30 = self.used_numbers
                    elif self.tower_type_high == "ZK_42":
                        self.used_numbers_single_ZK_42 = self.used_numbers

            if self.tower_type == "双回耐张塔":
                if self.tower_type_high == "SJ2_24" or self.tower_type_high == "SJ4_24":
                    self.used_numbers = round_up(self.double_circuit / 3, 0)
                    self.tower_number_weight = self.tower_weight * self.used_numbers
                    self.tower_area = self.used_numbers * self.tower_foot_distance ** 2

                    if self.tower_type_high == "SJ2_24":
                        self.used_numbers_double_SJ2_24 = self.used_numbers
                    elif self.tower_type_high == "SJ4_24":
                        self.used_numbers_double_SJ4_24 = self.used_numbers

            if self.tower_type == "双回直线塔":
                if self.tower_type_high == "SZ2_30" or self.tower_type_high == "SZK_42":
                    self.used_numbers = round_up(self.double_circuit * 1000 / 220 / 2, 0)
                    self.tower_number_weight = self.tower_weight * self.used_numbers
                    self.tower_area = self.used_numbers * self.tower_foot_distance ** 2

                    if self.tower_type_high == "SZ2_30":
                        self.used_numbers_double_SZ2_30 = self.used_numbers
                    elif self.tower_type_high == "SZK_42":
                        self.used_numbers_double_SZK_42 = self.used_numbers

            if self.tower_type == "铁塔电缆支架":
                if self.tower_type_high == "角钢":
                    self.used_numbers = self.tur_number
                    self.used_numbers_angle_steel = self.used_numbers
                    print(type(self.tower_weight),type(self.used_numbers))
                    self.tower_number_weight = self.tower_weight * self.used_numbers
                    self.tower_area = 0

        if self.project_chapter6_type == 0:
            if self.tower_type == "单回耐张塔":
                if self.tower_type_high == "J2_24" or self.tower_type_high == "J4_24":
                    self.used_numbers = round_up(self.single_circuit / 4, 0)
                    self.tower_number_weight = self.tower_weight * self.used_numbers
                    self.tower_area = self.used_numbers * self.tower_foot_distance ** 2
                    if self.tower_type_high == "J2_24":
                        self.used_numbers_single_J2_24 = self.used_numbers
                    elif self.tower_type_high == "J4_24":
                        self.used_numbers_single_J4_24 = self.used_numbers
                elif self.tower_type_high == "FS_18":
                    self.used_numbers = round_up(self.line_loop_number / 2, 0) + 1
                    self.tower_number_weight = self.tower_weight * self.used_numbers
                    self.tower_area = self.used_numbers * self.tower_foot_distance ** 2
                    self.used_numbers_single_FS_18 = self.used_numbers

            if self.tower_type == "单回直线塔":
                if self.tower_type_high == "Z2_30" or self.tower_type_high == "ZK_42":
                    if self.tower_type_high == "Z2_30":
                        self.used_numbers = round_up(self.single_circuit * 1000 * 0.8 / 260, 0)
                        self.used_numbers_single_Z2_30 = self.used_numbers
                    elif self.tower_type_high == "ZK_42":
                        self.used_numbers = round_up(self.single_circuit * 1000 * 0.2 / 290, 0)
                        self.used_numbers_single_ZK_42 = self.used_numbers
                    self.tower_number_weight = self.tower_weight * self.used_numbers
                    self.tower_area = self.used_numbers * self.tower_foot_distance ** 2
            if self.tower_type == "双回耐张塔":
                if self.tower_type_high == "SJ2_24" or self.tower_type_high == "SJ4_24":
                    self.used_numbers = round_up(self.double_circuit / 4, 0)
                    if self.tower_type_high == "SJ2_24":
                        self.used_numbers_double_SJ2_24 = self.used_numbers
                    elif self.tower_type_high == "SJ4_24":
                        self.used_numbers_double_SJ4_24 = self.used_numbers
                    self.tower_number_weight = self.tower_weight * self.used_numbers
                    self.tower_area = self.used_numbers * self.tower_foot_distance ** 2

            if self.tower_type == "双回直线塔":
                if self.tower_type_high == "SZ2_30" or self.tower_type_high == "SZK_42":
                    if self.tower_type_high == "SZ2_30":
                        self.used_numbers = round_up(self.double_circuit * 1000 * 0.6 / 280, 0)
                        self.used_numbers_double_SZ2_30 = self.used_numbers
                    elif self.tower_type_high == "SZK_42":
                        self.used_numbers = round_up(self.double_circuit * 1000 * 0.4 / 280, 0)
                        self.used_numbers_double_SZK_42 = self.used_numbers
                    self.tower_number_weight = self.tower_weight * self.used_numbers
                    self.tower_area = self.used_numbers * self.tower_foot_distance ** 2
            if self.tower_type == "铁塔电缆支架":
                if self.tower_type_high == "角钢":
                    self.used_numbers = self.tur_number
                    self.used_numbers_angle_steel = self.used_numbers
                    self.tower_number_weight = self.tower_weight * self.used_numbers
                    self.tower_area = 0

    def sum_cal_tower_type(self, tower_type_li, tower_type_high_li, tower_weight_li, tower_height_li,
                           tower_foot_distance_li):
        self.tower_type_list = tower_type_li
        self.tower_type_high_list = tower_type_high_li
        self.tower_weight_list = tower_weight_li
        self.tower_height_list = tower_height_li
        self.tower_foot_distance_list = tower_foot_distance_li

        for i in range(0, len(self.tower_weight_list)):

            TowerType.tower_type_models(self, self.tower_type_list[i], self.tower_type_high_list[i],
                                        self.tower_weight_list[i],
                                        self.tower_height_list[i], self.tower_foot_distance_list[i])
            # print(self.tower_type, self.tower_type_high, self.used_numbers)
            if self.tower_type == "铁塔电缆支架":
                self.sum_used_numbers = self.sum_used_numbers
            else:
                self.sum_used_numbers = int(self.used_numbers) + self.sum_used_numbers
            self.sum_tower_number_weight = float(self.tower_number_weight) + self.sum_tower_number_weight
            self.tower_number_area.append(self.tower_area)
            self.kilometer_tower_number = self.sum_used_numbers / (self.single_circuit + self.double_circuit)
            self.sum_tower_area = float(self.tower_area) + self.sum_tower_area


# tower_type_list = ['单回耐张塔', '单回耐张塔', '单回耐张塔', '单回直线塔', '单回直线塔', '双回耐张塔', '双回耐张塔', '双回直线塔', '双回直线塔', '铁塔电缆支架']
# tower_type_high_list = ['J2_24', 'J4_24', 'FS_18', 'Z2_30', 'ZK_42', 'SJ2_24', 'SJ4_24', 'SZ2_30', 'SZK_42', '角钢']
# tower_weight_list = [6.8, 8.5, 7, 5.5, 8.5, 12.5, 17, 6.5, 10, 0.5, ]
# tower_height_list = [32, 32, 27, 37, 49, 37, 37, 42, 54, 0]
# tower_foot_distance_list = [5.5, 5.5, 6, 5, 6, 7, 8, 6, 8, 0]
# project_chapter6_type = ['山地']
# project02 = TowerType(project_chapter6_type, 19, 22, 8, 1.5, 40, 6)
# project02.sum_cal_tower_type(tower_type_list, tower_type_high_list, tower_weight_list, tower_height_list,
#                              tower_foot_distance_list)
#
# print(project02.sum_tower_area, project02.sum_used_numbers, project02.sum_tower_number_weight,project02.kilometer_tower_number)
#
# project_chapter6_type = ['平地']
# tower_type_high_list = ['J2_24', 'J4_24', 'FS_18', 'Z2_30', 'ZK_42', 'SJ2_24', 'SJ4_24', 'SZ2_30', 'SZK_42', '角钢']
# tower_weight_list = [6.8, 8.5, 7, 5.5, 8.5, 12.5, 17, 6.5, 8.8, 0.5, ]
# tower_height_list = [32, 32, 27, 37, 49, 37, 37, 42, 54, 0]
# tower_foot_distance_list = [5.5, 5.5, 6, 5, 6, 7, 8, 6, 7, 0]
# project02 = TowerType(project_chapter6_type, 25.3, 23.6, 1.55, 3, 31, 5)
# project02.sum_cal_tower_type(tower_type_list, tower_type_high_list, tower_weight_list, tower_height_list,
#                              tower_foot_distance_list)
#
# print(project02.sum_tower_area, project02.sum_used_numbers, project02.sum_tower_number_weight,project02.kilometer_tower_number)
