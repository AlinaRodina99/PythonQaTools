# #!/usr/bin/env python3
# # vim: set ai et ts=4 sw=4:
#
# import matplotlib as mpl
# import matplotlib.pyplot as plt
# import math
#
# dpi = 80
# fig = plt.figure(dpi = dpi, figsize = (512 / dpi, 384 / dpi) )
# mpl.rcParams.update({'font.size': 10})
#
# plt.axis([0, 10, -1.5, 1.5])
#
# plt.title('Sine & Cosine')
# plt.xlabel('x')
# plt.ylabel('F(x)')
#
# xs = []
# sin_vals = []
# cos_vals = []
#
# x = 0.0
# while x < 10.0:
#     sin_vals += [ math.sin(x) ]
#     cos_vals += [ math.cos(x) ]
#     xs += [x]
#     x += 0.1
#
# plt.plot(xs, sin_vals, color = 'blue', linestyle = 'solid',
#          label = 'sin(x)')
# plt.plot(xs, cos_vals, color = 'red', linestyle = 'dashed',
#          label = 'cos(x)')
#
# plt.legend(loc = 'lower left')
# fig.savefig('trigan.png')

#!/usr/bin/env python3
# vim: set ai et ts=4 sw=4:

import matplotlib as mpl
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import datetime as dt
import csv

data_names = ['cafe', 'pharmacy', 'fuel', 'bank', 'waste_disposal',
              'atm', 'bench', 'parking', 'restaurant',
              'place_of_worship']
data_values = [9124, 8652, 7592, 7515, 7041, 6487, 6374, 6277,
               5092, 3629]

dpi = 80
fig = plt.figure(dpi = dpi, figsize = (1024 / dpi, 768 / dpi) )
mpl.rcParams.update({'font.size': 10})

plt.title('OpenStreetMap Point Types')

ax = plt.axes()
ax.yaxis.grid(True, zorder = 1)

xs = range(len(data_names))

plt.bar([x + 0.05 for x in xs], [ d * 0.9 for d in data_values],
        width = 0.2, color = 'red', alpha = 0.7, label = '2016',
        zorder = 2)
plt.bar([x + 0.3 for x in xs], data_values,
        width = 0.2, color = 'blue', alpha = 0.7, label = '2017',
        zorder = 2)
plt.xticks(xs, data_names)

fig.autofmt_xdate(rotation = 25)

plt.legend(loc='upper right')
fig.savefig('bars.png')