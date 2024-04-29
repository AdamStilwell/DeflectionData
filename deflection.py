import numpy as np
import csv
import math


def make_array_from_data(data, num):
    array = []
    for x in data:
        array.append(x[num])
    return array


def get_minimum(array):
    return np.min(array)


def get_maximum(array):
    return np.max(array)


def get_value_at_maximum(value_array, max_array):
    return value_array[np.argmax(max_array)]


def get_value_at_minimum(value_array, min_array):
    return value_array[np.argmin(min_array)]


def compile_select_data(*arrays):
    array = []
    for x in arrays:
        array.append(x)
    return array


class Deflection:
    def __init__(self, filename, headers_num):
        # will find a way to get these from the file headers
        self.headers = []
        with open(filename) as infile:
            reader = csv.reader(infile)
            for x in range(headers_num):
                self.headers.append(next(reader))

        # print(self.headers)
        self.sample_name = self.headers[1][1]
        self.weight = float(self.headers[2][1])
        self.test_force = float(self.headers[3][1])
        self.test_speed = float(self.headers[4][1])

        self.area = math.pi * (math.pow(0.0127, 2))
        self.pull_off_start = 0
        self.pull_off_ends = 0

        self.my_data = np.loadtxt(filename, delimiter=",", skiprows=headers_num, quotechar="\"")

        # make necessary arrays
        self.time_array = make_array_from_data(self.my_data, 0)
        self.sample_width_array = make_array_from_data(self.my_data, 1)
        self.sample_load_array = make_array_from_data(self.my_data, 2)
        self.pressure_array = self.make_pressure_array()
        self.psi_array = self.make_psi_array()
        self.h_delta_array = self.make_h_delta_array()
        self.time_step = self.time_array[2] - self.time_array[1]

        # pull off stuff
        self.find_pull_off()
        self.load_detach = get_value_at_minimum(self.sample_load_array, self.pressure_array)
        # need to talk to Jeff about how this is calculated
        # self.strain_to_break = self.sample_width_array[self.pull_off_ends] -
        # self.sample_width_array[self.pull_off_start]

        # sample size stuff
        self.minimum_gap = get_minimum(self.sample_width_array)
        self.width = self.sample_width_array[0]
        self.density = self.calculate_density()

        # deflection stuff
        self.deflection_array = self.make_deflection_array()
        self.max_deflection = get_maximum(self.deflection_array)

        # pressure stuff
        self.pressure_at_max_deflection = get_value_at_maximum(self.pressure_array, self.deflection_array)
        self.detach_pressure = get_minimum(self.pressure_array)
        self.stress_strain_array = self.make_stress_strain_array()

        self.full_data_array = self.compile_data()

    def make_pressure_array(self):
        array = []
        for x in self.sample_load_array:
            array.append((x/self.area)*0.000001)
        return array

    def make_psi_array(self):
        array = []
        for x in self.pressure_array:
            array.append(x*145.038)
        return array

    def make_deflection_array(self):
        array = []
        for x in self.sample_width_array:
            array.append(((self.width - x)/self.width) * 100)
        return array

    def make_stress_strain_array(self):
        array = []
        for i in range(len(self.deflection_array)):
            array.append(self.pressure_array[i] * 1000000 * self.deflection_array[i] / 100)
        return array

    def make_h_delta_array(self):
        array = []
        for x in self.sample_width_array:
            array.append(self.test_speed/x/1000)
        return array

    def find_pull_off(self):
        i = 0
        for x in self.sample_load_array:
            if x < 0 and self.pull_off_start == 0:
                self.pull_off_start = i
                i += 1
            elif x > 0 and self.pull_off_start > 0:
                self.pull_off_ends = i-2
                break
            i += 1

    def calculate_density(self):
        return self.weight/(self.area * self.width) * 0.001

    def compile_data(self):
        array = [self.time_array, self.sample_width_array, self.sample_load_array, self.deflection_array,
                 self.pressure_array, self.psi_array, self.stress_strain_array, self.h_delta_array]
        self.headers[-2].append("Deflection")
        self.headers[-2].append("Pressure")
        self.headers[-2].append("PSI")
        self.headers[-2].append("Stress * strain")
        self.headers[-2].append("h_dot/h")
        return array
