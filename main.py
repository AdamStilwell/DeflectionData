import numpy as np
import pandas as pd


class Deflection():
    def __init__(self, filename):
        # will find a way to get these from the file headers
        self.weight = 2.924
        self.area = 0.000506707
        self.pull_off_start = 0
        self.pull_off_ends = 0
        self.test_speed = 5.0
        self.time_step = 0.1

        self.my_data = np.loadtxt(filename, delimiter=",", skiprows=11, quotechar="\"")

        self.sample_width_array = self.make_width_array()
        self.sample_load_array = self.make_load_array()
        self.pressure_array = self.make_pressure_array()

        self.find_pull_off()

        self.width = self.get_sample_width()
        self.deflection_array = self.make_deflection_array()

        self.detach_pressure = self.get_detach_pressure()
        self.density = self.calculate_density()

    def make_width_array(self):
        array = []
        for x in self.my_data:
            array.append(x[1])
        return array

    def make_load_array(self):
        array = []
        for x in self.my_data:
            array.append(x[2])
        return array

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

    def get_detach_pressure(self):
        return np.min(self.pressure_array)

    def get_sample_width(self):
        return self.sample_width_array[0]

    def calculate_density(self):
        return self.weight/(self.area * self.width) * 0.001


if __name__ == "__main__":
    my_deflection = Deflection(filename="Specimen_RawData_1.csv")
    # print out data to excel sheet here
    print("test")

