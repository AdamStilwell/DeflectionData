import numpy as np
import csv
import math
from scipy.optimize import curve_fit


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


def func_final(x, a, n, b):
    return a * (x - b) ** n


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

        self.my_data = np.loadtxt(filename, delimiter=",", skiprows=headers_num, quotechar="\"")

        # make necessary arrays
        self.time_array = make_array_from_data(self.my_data, 0)
        self.sample_width_array = make_array_from_data(self.my_data, 1)
        self.sample_load_array = make_array_from_data(self.my_data, 2)
        self.my_data = None
        self.pressure_array = self.make_pressure_array()
        self.psi_array = self.make_psi_array()
        self.h_delta_array = self.make_h_delta_array()
        self.time_step = self.time_array[2] - self.time_array[1]

        # sample size stuff
        self.minimum_gap = get_minimum(self.sample_width_array)
        self.width = self.find_sample_width()
        self.density = self.calculate_density()
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
        self.pull_off_ends = self.find_pull_off_end()
        self.load_detach = get_value_at_minimum(self.sample_load_array, self.pressure_array)
        # print(self.time_array[self.pull_off_start])
        # print(self.time_array[self.pull_off_ends])
        self.strain_to_break = (self.sample_width_array[self.pull_off_ends] -
                                self.sample_width_array[self.pull_off_start])
        self.g1c = self.calculate_g1c()

        # power law stuff
        self.offset = self.test_speed / self.width / 1000
        self.power_law_first = []
        self.perr_popt_pcov_dict = {}
        self.power_law_values = self.power_law_calculation()

        self.full_data_array = self.compile_data()

    def func_first(self, x, a, n):
        return a * (x - self.offset) ** n

    def func(self, x, b):
        return self.power_law_first[0] * (x - b) ** self.power_law_first[1]

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

    def make_h_delta_array(self):
        array = []
        for x in self.sample_width_array:
            array.append(self.test_speed / x / 1000)
        return array

    def find_pull_off_start(self):
        i = np.argmax(self.sample_load_array)
        for x in range(i, len(self.stress_strain_array)):
            if self.stress_strain_array[i] < 0:
                return i
            else:
                i += 1

    def find_pull_off_end(self):
        target = self.detach_pressure / 10
        i = np.argmin(self.pressure_array) + 1
        pull_off = 0
        min_diff = abs(target - self.pressure_array[i])
        for x in range((np.argmin(self.pressure_array) + 1), len(self.pressure_array)):
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
        for x in range(300):
            if self.sample_load_array[i] > 2:
                break
            else:
                diff = abs(self.sample_load_array[i] - target)
                if diff < min_diff:
                    min_diff = diff
                    width = i
                i += 1
        return self.sample_width_array[width]

    def calculate_density(self):
        return self.weight / (self.area * self.width) * 0.001

    def calculate_g1c(self):
        return (sum(self.stress_strain_array[self.pull_off_start:self.pull_off_ends])
                * (self.test_speed / 1000000)
                * self.time_step
                * -1)

    def curve_fit_func(self, start, end):
        self.power_law_first = curve_fit(self.func_first,
                                         self.h_delta_array[start:end],
                                         self.pressure_array[start:end],
                                         p0=(100, 1),
                                         maxfev=1000,
                                         nan_policy="omit")[0]
        power_law_offset = curve_fit(self.func,
                                     self.h_delta_array[start:end],
                                     self.pressure_array[start:end],
                                     p0=self.offset,
                                     maxfev=1000,
                                     nan_policy="omit")[0]
        try:
            power_law_popt_pcov = curve_fit(func_final,
                                            self.h_delta_array[start:end],
                                            self.pressure_array[start:end],
                                            p0=(self.power_law_first[0], self.power_law_first[1],
                                                power_law_offset[0]),
                                            maxfev=10000,
                                            nan_policy="omit")
        except RuntimeError:
            return -1
        value = [make_array_from_numpy_array(power_law_popt_pcov, 0),
                 make_array_from_numpy_array(power_law_popt_pcov, 1)]
        perr = np.sqrt(np.diag(power_law_popt_pcov[1]))
        perr_avg = np.average(perr)
        self.perr_popt_pcov_dict.update({perr_avg: value})
        return perr_avg

    def power_law_calculation(self):
        # pick a start point by the power of DEDUCTION
        end = np.argmax(self.sample_load_array)
        start = end - int(end / 7)
        self.curve_fit_func(start=start,
                            end=end)
        start_min = int(len(self.deflection_array) / 10)
        # loop the curve fit
        while True:
            fit = True
            start = int(start / 4)
            if start < start_min:
                start = start_min
                fit = False

            perr_avg = self.curve_fit_func(start=start,
                                           end=end)
            if perr_avg == -1:
                break
            # when perr is greater than the perr of the previous curves end the loop
            for x in self.perr_popt_pcov_dict.keys():
                if perr_avg > x:
                    fit = False
                    break
            if not fit:
                break
        min_key = min(self.perr_popt_pcov_dict.keys())
        return self.perr_popt_pcov_dict.get(min_key)

    def compile_data(self):
        array = [self.time_array, self.sample_width_array, self.sample_load_array, self.deflection_array,
                 self.pressure_array, self.psi_array, self.stress_strain_array, self.h_delta_array]
        more_headers = ["Deflection", "Pressure", "PSI", "Stress * strain", "h_dot/h"]
        for x in more_headers:
            self.headers[-2].append(x)
        return array
