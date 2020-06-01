import matplotlib.pyplot as plt
# 使得matplotlib可以正确显示中文
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False
x = y = path = width = item = 0
# ------------------------------------------------------------------------------------------------------------------
# 图片总体大小
plt.figure(figsize=(10, 3))
# 直线图
plt.plot(x, y, ls='-', c='k')
# 散点图
plt.scatter(x, y, marker='.',c='b')
cm = plt.cm.get_cmap('RdYlBu_r')
plt.scatter(x_value_all, y_value_all, c=z_value_all, vmin=50, vmax=100, marker='.', cmap=cm)
# 柱状图
plt.bar(x, y, width=width, label=item, bottom=value)
# 设置坐标轴上下限
plt.xlim(xmax=0, xmin=100)
plt.ylim(ymax=140, ymin=0)
ax2 = plt.twinx()
ax2.set_xlim(xmax=100, xmin=0)
# 设置坐标轴的文字标签
plt.xlabel('Time')
plt.ylabel('Speed')
# 设置坐标轴零点对齐
ax = plt.gca()
ax.spines['bottom'].set_position(('data', 0))
ax.spines['left'].set_position(('data', 0))
# 辅助线（竖线）
plt.axvline(x=10, c='b', ls='dashed')
# 辅助线（横线）
plt.hlines(y=80, xmin=0, xmax=100, colors='r', linestyles='dashed')
# 在图上添加独立的文字
for x, y in zip(x_value, y_value_all):
    plt.text(x=x, y=min(y * 1.05, y + 1), s='%.0f' % y, ha='center', va='bottom', fontsize=10)
# 填充色
plt.fill_between(x, y1, y2, facecolor='b')
# 保存图片
plt.savefig(path, bbox_inches='tight')
plt.close('all')
# 坐标轴刻度的文字
x_axis_info = []
plt.xticks(x, x_axis_info)
# 图例
plt.legend(labels=car_number_list, loc='upper center')
plt.legend(labels=car_number_list, edgecolor='white', ncol=min(len(car_number_list), 8),
           bbox_to_anchor=(0, 1.02 + 0.04 * int(len(car_number_list) / 8), 1, .102),
           loc='upper left', mode='expand', prop={'size': 20})
# 显示图片（会使程序暂停直到图片关闭）
plt.show()

# 绘制等高线图，并用散点图标示数据点
def draw_pic_contour():
    for i_file, file_name in enumerate(os.listdir(pic_contour_excel_path)):
        wbk = openpyxl.load_workbook(pic_contour_excel_path + '\\' + file_name)
        sht = wbk.worksheets[0]
        signal_list = ['']
        x_start = 3
        x_name = x_contour_name
        y_name = y_contour_name
        y = 1
        while sht.cell(row=1, column=y).value != None:
            signal_list.append(sht.cell(row=1, column=y).value)
            y += 1
        y_x_name = signal_list.index(x_name)
        y_y_name = signal_list.index(y_name)
        for i_signal, signal_name in enumerate(pic_contour_signal_list):
            print(signal_name)
            x_value = []
            y_value = []
            z_value = []
            pic_path = os.getcwd() + '/03_Result/Scatter/%s.png' % (signal_name)
            if signal_name in signal_list:
                y = signal_list.index(signal_name)
            else:
                continue
            x = 0
            # 依次录入各组数据
            while not (sht.cell(row=x_start + x, column=y).value == None and sht.cell(row=x_start + x + 1, column=y).value == None):
                if sht.cell(row=x_start + x, column=y).value != None:
                    x_value.append(sht.cell(row=x_start + x, column=y_x_name).value)
                    y_value.append(sht.cell(row=x_start + x, column=y_y_name).value)
                    z_value.append(sht.cell(row=x_start + x, column=y).value)

                x += 1
            plt.figure(figsize=(15, 10))
            plt.scatter(x_value, y_value, marker='.', color='b', label=z_value)
            # 为数据点添加标签
            for a, b, c in zip(x_value, y_value, z_value):
                plt.text(a, b, c, fontdict=None, withdash=False, ha='center', va='bottom')
            # 等高线
            # for a, b, c in zip(x_value, y_value, z_value):
            #     if c >
            #     plt.plot(x, y, '-bo')
            # plt.show()
            x = 0
            x_value = []
            y_value = []
            z_value = []
            x_temp = []
            y_temp = []
            z_temp = []
            count = 0
            while not (sht.cell(row=x_start + x, column=y).value == None and sht.cell(row=x_start + x + 1, column=y).value == None):
                if sht.cell(row=x_start + x, column=y).value != None:
                    count += 1
                    x_temp.append(sht.cell(row=x_start + x, column=y_x_name).value)
                    y_temp.append(sht.cell(row=x_start + x, column=y_y_name).value)
                    z_temp.append(sht.cell(row=x_start + x, column=y).value)
                    # z_value = np.append(z_value, sht.cell(row=x_start + x, column=y).value)
                else:
                    count = 0
                    x_value.append(sht.cell(row=x_start + x, column=y_x_name).value)
                    y_value.append(y_temp)
                    z_value.append(z_temp)
                    x_temp = []
                    y_temp = []
                    z_temp = []
                x += 1
            y_value.append(y_temp)
            z_value.append(z_temp)

            X = [[1000, 1200, 1600, 2000, 2400, 2800, 3200, 3600, 4000, 4400, 4800, 5200, 5600, 6000, 6500] for x in range(10)]
            # 去除中间多余的点使得y矩阵为方形
            length_min = len(y_value[0])
            for i, item in enumerate(y_value):
                if len(item) < length_min:
                    length_min = len(item)
            y_value_new = []
            z_value_new = []
            for i, item in enumerate(y_value):
                if len(item) > length_min:
                    y_temp = [0 for x in range(length_min)]
                    z_temp = [0 for x in range(length_min)]
                    len_first_half = int(length_min / 2)
                    len_last_half = len(item) - len_first_half
                    for i_temp, item_temp in enumerate(item[:len_first_half]):
                        y_temp[i_temp] = item_temp
                        z_temp[i_temp] = z_value[i][i_temp]
                    for i_temp, item_temp in enumerate(item[len_last_half:]):
                        y_temp[i_temp + len_first_half] = item_temp
                        z_temp[i_temp + len_first_half] = z_value[i][i_temp + len_last_half]
                    y_value_new.append(y_temp)
                    z_value_new.append(z_temp)
                else:
                    y_value_new.append(item)
                    z_value_new.append(z_value[i])
            Y = np.transpose(y_value_new)
            # Z = z_value.reshape(len(X[0]),len(y_value[0]))
            Z = np.transpose(z_value_new)
            levels = [0, 20, 40, 60, 80]
            # 填充颜色
            plt.contourf(X, Y, Z, levels=levels, alpha = 0.6, colors = ['#00FF00', '#7DFF00', '#FF7D00', '#FF0000'])
            # 绘制等高线
            C = plt.contour(X, Y, Z, levels=levels, colors='black', linewidth=0.5)
            # plt.scatter(X, Y, marker='.', c='k')
            # 显示各等高线的数据标签
            plt.clabel(C, inline = True, fontsize = 10)
            ax = plt.gca()
            ax.spines['bottom'].set_position(('data', 0))
            ax.spines['left'].set_position(('data', 0))

            plt.savefig(pic_path, bbox_inches='tight')
            plt.close('all')

# 多个子图
def analysis_trip_overvie(self):
        print('start drawing pics')
        # 设置图片大小
        plt.figure(figsize=self.figsize)
        p_height = 4    # 大图的大小是小图的4倍
        pic_height_total = p_height + len(self.emission_name) * p_height + len(self.signal_list_2)
        gs = GridSpec(pic_height_total, 1)
        # ----------------------------------------------------------------------------------------------------------------------
        # 车速
        # ----------------------------------------------------------------------------------------------------------------------
        plt.subplot(gs[0:p_height, 0])
        speed_value = self.data_file.getChannelData(self.speed_name)[self.t_start_inca: self.t_end_inca]
        plt.plot(self.time_value, speed_value, ls='-', c='b', linewidth=1)
        plt.ylabel('Speed (km/h)', size=self.label_size, rotation='horizontal', horizontalalignment='right', verticalalignment='center')
        plt.xticks([])
        plt.tick_params(labelsize=self.tick_label_size)
        ax = plt.gca()
        ax.spines['bottom'].set_position(('data', 0))
        ax.spines['left'].set_position(('data', 0))
        plt.xlim(xmax=max(self.time_value), xmin=0)
        # ----------------------------------------------------------------------------------------------------------------------
        # 排放报告中的排放物（NOx、CO、PN）
        # ----------------------------------------------------------------------------------------------------------------------
        i = 1
        for emission_name in self.emission_name:
            print(emission_name)
            i_emission = self.emission_values_name.index(emission_name)
            emission_value = self.emission_values[i_emission][self.t_start_emission: self.t_end_emission]
            # p_subplot = int(p_subplot_1 + '1' + str(i + 1))
            # plt.subplot(p_subplot)
            plt.subplot(gs[i * p_height:i * p_height + p_height, 0])
            plt.plot(self.time_value, emission_value, ls='-', c='b', linewidth=1)
            text_ylabel = '%s (%s)' % (self.emission_values_name[i_emission], self.emission_values_unit[i_emission])
            plt.ylabel(text_ylabel, size=self.label_size, rotation='horizontal', horizontalalignment='right', verticalalignment='center')
            plt.xticks([])
            plt.tick_params(labelsize=self.tick_label_size)
            ax = plt.gca()
            ax.spines['bottom'].set_position(('data', 0))
            ax.spines['left'].set_position(('data', 0))
            plt.xlim(xmax=max(self.time_value), xmin=0)
            i += 1
        # ----------------------------------------------------------------------------------------------------------------------
        # INCA文件中的变量（Cat heating等）
        # ----------------------------------------------------------------------------------------------------------------------
        n_temp = i
        for i_signal, signal_name in enumerate(self.signal_list_2):
            print(self.signal_name_list_2[i_signal])
            if signal_name:
                signal_value = self.data_file.getChannelData(signal_name)[self.t_start_inca: self.t_end_inca]
                name_new = self.signal_name_list_2[i_signal]
            else:
                signal_value = np.zeros(speed_value.shape)
                name_new = self.signal_name_list_2[i_signal] + '\n(not found)'
            plt.subplot(gs[n_temp * (p_height - 1) + i, 0])
            plt.plot(self.time_value, signal_value, ls='-', c='b', linewidth=1)
            plt.fill_between(self.time_value, signal_value, np.zeros(speed_value.shape), facecolor='b')
            plt.ylabel(name_new, size=self.label_size, rotation='horizontal', horizontalalignment='right', verticalalignment='center')
            if i_signal < len(self.signal_list_2) - 1:
                plt.xticks([])
            else:
                plt.xlabel('Time (s)', size=self.label_size)
            plt.yticks([])
            plt.tick_params(labelsize=self.tick_label_size)
            plt.ylim(ymax=1.5, ymin=0)
            ax = plt.gca()
            ax.spines['bottom'].set_position(('data', 0))
            ax.spines['left'].set_position(('data', 0))
            plt.xlim(xmax=max(self.time_value), xmin=0)
            i += 1
        # 保存并插入图片
        pic_path = temp_folder + 'general.jpg'
        plt.savefig(pic_path, bbox_inches='tight')
        plt.close()
        slide = self.ppt.slides[1]
        slide.shapes[0].text_frame.paragraphs[0].text = 'Trip Overvie'
        slide.shapes.add_picture(image_file=pic_path, left=self.pic_left, top=self.pic_top, width=self.pic_width,
                                 height=self.pic_height)