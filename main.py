import numpy as np
import pandas as pd


def make_array_from_data(data, num):
    array = []
    for x in data:
        array.append(x[num])
    return array


def get_minimum(array):
    return np.min(array)


def get_maximum(array):
    return np.max(array)


def get_value_at_maximum(array1, array2):
    return array1[np.argmax(array2)]


def get_value_at_minimum(array1, array2):
    return array1[np.argmin(array2)]


class Deflection:
    def __init__(self, filename):
        # will find a way to get these from the file headers
        self.weight = 2.924
        self.area = 0.000506707
        self.pull_off_start = 0
        self.pull_off_ends = 0
        self.test_speed = 5.0
        self.time_step = 0.1

        self.my_data = np.loadtxt(filename, delimiter=",", skiprows=11, quotechar="\"")

        # make necessary arrays
        self.sample_width_array = make_array_from_data(self.my_data, 1)
        self.sample_load_array = make_array_from_data(self.my_data, 2)
        self.pressure_array = self.make_pressure_array()

        self.find_pull_off()
        self.strain_to_break = self.sample_width_array[self.pull_off_ends] - self.sample_width_array[self.pull_off_start]

        self.minimum_gap = get_minimum(self.sample_width_array)

        self.width = self.sample_width_array[0]
        self.deflection_array = self.make_deflection_array()
        self.max_deflection = get_maximum(self.deflection_array)
        self.pressure_at_max_deflection = get_value_at_maximum(self.pressure_array, self.deflection_array)

        self.detach_pressure = get_minimum(self.pressure_array)
        self.density = self.calculate_density()

        self.full_data_array = self.compile_data()

    def make_pressure_array(self):
        array = []
        for x in self.sample_load_array:
            array.append((x/self.area)*0.000001)
        return array

    def make_deflection_array(self):
        array = []
        for x in self.sample_width_array:
            array.append(((self.width - x)/self.width) * 100)
        return array

    def find_pull_off(self):
        i = 0
        for x in self.sample_load_array:
            if x < 0 and self.pull_off_start == 0:
                self.pull_off_start = i
                i += 1
            elif x > 0 and self.pull_off_start > 0:
                self.pull_off_ends = i
                break
            i += 1

    def calculate_density(self):
        return self.weight/(self.area * self.width) * 0.001

    def compile_data(self):
        array = [self.sample_width_array, self.sample_load_array, self.deflection_array, self.pressure_array]
        return array


if __name__ == "__main__":
    # file_name = file_name.split("/")[-1]
    my_deflection = Deflection(filename="Specimen_RawData_1.csv")
    # print out data to Excel sheet here
    print(my_deflection.max_deflection)
