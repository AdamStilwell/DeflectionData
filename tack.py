import numpy as np
import csv
import math


def make_array_from_data(data, num):
    array = []
    for x in data:
        array.append(x[num])
    return array


def make_array_from_numpy_array(numpy_array, sub_array):
    array = []
    for x in numpy_array[sub_array]:
        array.append(x)
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


class Tack:
    def __init__(self, filename, headers_num, headers_offset):
        # will find a way to get these from the file headers
        self.headers = []
        with open(filename) as infile:
            reader = csv.reader(infile)
            for x in range(headers_num + headers_offset):
                self.headers.append(next(reader))

        # print(self.headers)
        self.sample_name = self.headers[1 + headers_offset][1]
        self.test_force = float(self.headers[3 + headers_offset][1])
        self.test_speed = float(self.headers[4 + headers_offset][1])

        self.area = math.pi * (math.pow(0.0127, 2))

        self.my_data = np.loadtxt(filename, delimiter=",", skiprows=headers_num + headers_offset, quotechar="\"")

        # print(self.my_data)
        # make necessary arrays
        self.time_array = make_array_from_data(self.my_data, 0)
        self.sample_width_array = make_array_from_data(self.my_data, 1)
        self.sample_load_array = make_array_from_data(self.my_data, 2)
        self.my_data = None
        self.pressure_array = self.make_pressure_array()
        self.psi_array = self.make_psi_array()
        self.time_step = self.time_array[2] - self.time_array[1]

        # sample size stuff
        self.minimum_gap = get_minimum(self.sample_width_array)
        self.width = self.find_sample_width()
        self.max_load = get_maximum(self.sample_load_array)

        # deflection stuff
        self.deflection_array = self.make_deflection_array()
        self.max_deflection = get_maximum(self.deflection_array)

        # pressure stuff
        self.pressure_at_max_deflection = get_value_at_maximum(self.pressure_array, self.deflection_array)
        self.detach_pressure = get_minimum(self.pressure_array)
        self.stress_strain_array = self.make_stress_strain_array()

        # pull off stuff
        self.pull_off_start = self.find_pull_off_start()
        self.pull_off_final = self.find_pull_off_final()
        self.pull_off_ends = self.find_pull_off_end()
        self.load_detach = get_value_at_minimum(self.sample_load_array, self.pressure_array)
        # print(self.time_array[self.pull_off_start])
        # print(self.time_array[self.pull_off_ends])
        self.strain_to_break = (self.sample_width_array[self.pull_off_ends] -
                                self.sample_width_array[self.pull_off_start])
        self.g1c = self.calculate_g1c()

        self.full_data_array = self.compile_data()

    def make_pressure_array(self):
        array = []
        for x in self.sample_load_array:
            array.append((x / self.area) * 0.000001)
        return array

    def make_psi_array(self):
        array = []
        for x in self.pressure_array:
            array.append(x * 145.038)
        return array

    def make_deflection_array(self):
        array = []
        for x in self.sample_width_array:
            array.append(((self.width - x) / self.width) * 100)
        return array

    def make_stress_strain_array(self):
        array = []
        for i in range(len(self.deflection_array)):
            array.append(self.pressure_array[i] * 1000000 * self.deflection_array[i] / 100)
        return array

    def find_pull_off_start(self):
        i = np.argmax(self.sample_load_array)
        for x in range(i, len(self.stress_strain_array)):
            if self.stress_strain_array[i] < 0:
                return i
            else:
                i += 1
        if i == len(self.stress_strain_array):
            return i - 1

    def find_psi_values(self, psi):
        target = psi
        avg_array = []
        i = 0
        values = False
        for x in range(len(self.psi_array)):
            if self.psi_array[i] > (psi + 0.5):
                break
            else:
                if self.psi_array[i] >= target:
                    avg_array.append(self.deflection_array[i])
                    values = True
                i += 1
        if not values:
            return 0
        return np.average(avg_array)

    def find_pull_off_final(self):
        for x in range(self.pull_off_start, len(self.pressure_array)):
            if self.stress_strain_array[x] > 0:
                return x
            else:
                return len(self.pressure_array)

    # this function determines the position in the array closest to 1/10th of the max pressure
    def find_pull_off_end(self):
        target = self.detach_pressure / 10
        i = np.argmin(self.pressure_array) + 1
        if i == len(self.pressure_array):
            i -= 1
        pull_off = 0
        min_diff = abs(target - self.pressure_array[i])
        for x in range((np.argmin(self.pressure_array) + 1), self.pull_off_final):
            diff = abs(self.pressure_array[x] - target)
            if diff < min_diff:
                min_diff = diff
                pull_off = x
        return pull_off

    def find_sample_width(self):
        target = 1
        i = 0
        min_diff = abs(target - self.sample_load_array[i])
        width = 0
        for x in range(len(self.sample_load_array)):
            if self.sample_load_array[i] > 2:
                break
            else:
                diff = abs(self.sample_load_array[i] - target)
                if diff < min_diff:
                    min_diff = diff
                    width = i
                i += 1
        return self.sample_width_array[width]

    def calculate_g1c(self):
        return (sum(self.stress_strain_array[self.pull_off_start:self.pull_off_ends])
                * (self.test_speed / 1000000)
                * self.time_step
                * -1)

    def compile_data(self):
        array = [self.time_array, self.sample_width_array, self.sample_load_array, self.deflection_array,
                 self.pressure_array, self.psi_array, self.stress_strain_array, self.h_delta_array]
        more_headers = ["Tack", "Pressure", "PSI", "Stress * strain", "h_dot/h"]
        for x in more_headers:
            self.headers[-2].append(x)
        return array