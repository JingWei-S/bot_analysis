import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import os

path = "/Users/user/Desktop/chatbot_online"

data = pd.read_excel(os.path.join(path, 'sample_coded_data.xlsx'))
data = data[data['Speaker'] == 'system']
loop_columns = ['Acknowledging', 'Appreciating', 'Thanking', 'Sympathizing', 'Advice/...', 'Reject/...']
labels = ['Acknowledging', 'Appreciating', 'Thanking', 'Sympathizing', 'Advice', 'Reject']
color_map = ['Greys_r', 'Purples_r', 'Blues_r', 'Greens_r', 'BuGn_r', 'cividis']


def print_plot(data, color_map, m, n):
    count_mat = np.zeros((len(loop_columns), len(loop_columns)))
    for i in range(len(loop_columns)):
        temp_data = data[loop_columns[i:]]
        temp_data_size = len(temp_data)
        # print(i)
        # print('current: ' + loop_columns[i])
        temp_data.dropna(subset=[loop_columns[i]], inplace=True)
        # temp_data.head()
        for j in range(len(temp_data.columns)):
            count_mat[i, i + j] = (temp_data.iloc[:, j].sum()) / temp_data_size * 100
    print(count_mat)

    for i in range(len(count_mat)):
        x = np.ones(len(count_mat)) * i * 2 + m
        y = np.array(list(range(len(count_mat)))) * 2 + np.ones(len(count_mat)) * n
        s = count_mat[i, :]
        plt.scatter(x, y, s=s, cmap=color_map, c=y, label='test')
        # ax = plt.gca()
        # plt.legend(['Condition 1'])

    # plt.show()


print_plot(data[data.Condition == 1], 'Greys_r', -0.25, -0.25)

print_plot(data[data.Condition == 2], 'Purples_r', -0.25, 0.25)

print_plot(data[data.Condition == 3], 'Blues_r', 0.25, -0.25)

print_plot(data[data.Condition == 4], 'Greens_r', 0.25, 0.25)

plt.xticks(np.linspace(0, 10, 6), labels, rotation=45)
plt.yticks(np.linspace(0, 10, 6), labels)
fig = plt.gcf()
fig.set_size_inches(10, 10)
# plt.legend(loc=0)
plt.show()

empathy_count = np.zeros((4, len(loop_columns)))
x_data = np.arange(len(loop_columns)-1)
width = 0.125
diff_ = [-width*1.5, -width/2, width/2, width*1.5]
for i in [1, 2, 3, 4]:
    temp_data = data[data.Condition == i]
    temp_data_len = len(temp_data)
    for j in range(len(loop_columns)-1):
        empathy_count[int(i) - 1, j] = temp_data[loop_columns[j]].dropna().count() / temp_data_len * 100

    plt.bar(x_data + diff_[int(i)-1], empathy_count[int(i) - 1,:-1], width=width, label = i)
plt.xticks(range(5), loop_columns[:-1])
plt.legend(loc=0)
plt.show()

essential_data = pd.read_excel(os.path.join(path, 'essential_coded_data.xlsx'))
essential_columns = ['Greeting', 'Topic opening', 'RQ', 'RA', 'OQ', 'OA', 'Statement', 'Closing']
component_mat = np.zeros((4, 1 + len(essential_columns)))
for i in [1, 2, 3, 4]:
    temp_data = essential_data[essential_data.Condition == i]
    for j in range(len(essential_columns)):
        component_mat[int(i) - 1, j+1] = temp_data[essential_columns[j]].dropna().count()

for i in range(4):
    component_mat[i,:] = component_mat[i,:]/sum(component_mat[i,:])

for i in range(1,8):
    plt.bar(['Condition 1', 'Condition 2', 'Condition 3', 'Condition 4'], component_mat[:,i], bottom=component_mat[:,:i].sum(axis=1), label = essential_columns[i-1])
plt.legend(loc=0, bbox_to_anchor=(1, 0.5))
fig = plt.gcf()
fig.set_size_inches(12, 8)
plt.show()



print('Hello!')
