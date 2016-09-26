import docx
import numpy as np
from docx import Document
from bokeh.models import HoverTool, BoxSelectTool
from bokeh.models.widgets import Panel, Tabs
from bokeh.io import output_file, show
from bokeh.plotting import figure
from bokeh.models.formatters import NumeralTickFormatter
from bokeh.models import FixedTicker



data = []
ys = []
xs = []
document = Document('Comedy_Central_Questionnaire.docx')

for p in document.paragraphs:
    data.append(p.text)

data = data[13:]
data[6] = "2014\tWeek - 7\t3393847.333"
num = 0
for info in data:
    data[num] = info.split("\t")
    num+=1

for i in data:
    for item in i:
       if 'O' in item:
           i[2] = item = item.replace('O','0')
           print item
    ys.append(float(i[2]))
    xs.append(i[0]+i[1])
ys1 = ys[:20]
ys2 = ys[20:]


output_file("slider.html")


TOOLS1 = [BoxSelectTool(), HoverTool(tooltips = [
    ("index", "$index"),
    ("(x,y)", "($x, $y)"),
    ("fill color", "$color[swatch]:color"),
])]
TOOLS2 = [BoxSelectTool(), HoverTool(tooltips = [
    ("index", "$index"),
    ("(x,y)", "($x, $y)"),
    ("fill color", "$color[swatch]:color"),
])]
TOOLS3 = [BoxSelectTool(), HoverTool(tooltips = [
    ("index", "$index"),
    ("(x,y)", "($x, $y)"),
    ("line color", "$color[swatch]:line_color"),
])]


p1 = figure(plot_width=600, plot_height=600, title="2014 - Season's New Viewers", tools=TOOLS1)
p1.line(np.arange(1,21), ys1, line_width=3, line_color="red", alpha=0.7)
p1.yaxis.formatter = NumeralTickFormatter(format="0,0")
tab1 = Panel(child=p1, title="2014 Season")

p1.yaxis.axis_label = "New Viewers"
p1.yaxis.major_label_orientation = "vertical"

p1.xaxis.axis_label = "Week"
p1.xaxis[0].ticker=FixedTicker(ticks=np.arange(1,21))

p2 = figure(plot_width=600, plot_height=600, title="2015 - Season's New Viewers", tools=TOOLS2)
p2.line(np.arange(1,21), ys2, line_width=3, line_color="green", alpha=0.5)
p2.yaxis.formatter = NumeralTickFormatter(format="0,0")
tab2 = Panel(child=p2, title="2015 Season")

p2.yaxis.axis_label = "New Viewers"
p2.yaxis.major_label_orientation = "vertical"

p2.xaxis.axis_label = "Week"
p2.xaxis[0].ticker=FixedTicker(ticks=np.arange(1,21))

p3 = figure(plot_width=600, plot_height=600, title="2014&2015 - Season's New Viewers", tools=TOOLS3)
p3.line(np.arange(1,21), ys2, line_width=3, legend="2015", line_color="green", alpha=0.5)
p3.line(np.arange(1,21), ys1, line_width=3, legend="2014", line_color="red", alpha=0.7)
p3.yaxis.formatter = NumeralTickFormatter(format="0,0")
tab3 = Panel(child=p3, title="Both Seasons")
p3.legend.location = "top_left"

p3.yaxis.axis_label = "New Viewers"
p3.yaxis.major_label_orientation = "vertical"

p3.xaxis.axis_label = "Week"
p3.xaxis[0].ticker=FixedTicker(ticks=np.arange(1,21))

tabs = Tabs(tabs=[ tab1, tab2, tab3 ])

show(tabs)
