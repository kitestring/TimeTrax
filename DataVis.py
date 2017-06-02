import matplotlib.pyplot as plt
plt.rcdefaults()
import numpy as np
from pylab import *

class DataVisualizer():
	def create_horizontal_bar_chart(self, data_labels, data_values, bar_color, x_axis_label, chart_title):
		y_pos = np.arange(len(data_labels))
		plt.barh(y_pos, data_values, align='center', color=bar_color)
		plt.yticks(y_pos, data_labels)
		plt.xlabel(x_axis_label)
		plt.title(chart_title)
		plt.grid(True)
		plt.savefig('temp_bar_chart.png', bbox_inches='tight')
		plt.clf()

	def create_pie_chart(self, labels, data_values, chart_title):
		#This works, but it's ugly as shit
		percents = self.calc_percentages(data_values)
		figure(1, figsize=(10,10))
		pie(percents,labels=labels, autopct='%1.1f%%')
		savefig('bar.png')
		clf()
		
				
	def calc_percentages(self, data_values):
		percentages = []
		summation = sum(data_values)
		for value in data_values:
			percent = (value/summation) * 100.00
			percentages.append(percent)
		return percentages