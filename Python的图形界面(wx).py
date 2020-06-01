
import pptx, wx, openpyxl
# ----------------------------------------------------------------------------------------------------------------------
# 主程序
# ----------------------------------------------------------------------------------------------------------------------
class MainWindow(wx.Frame):

    #print(code_value)
    #print(file_number_list)
    # 窗口初始化
    # ------------------------------------------------------------------------------------------------------------------
    def __init__(self,parent,title):
        global code_keylist_en, code_value, keyword_count, code_keyword
        self.i_number = code_keylist_en.index('Nr')   # 编号在字段中的位置
        self.i_attachment = code_keylist_en.index('Attachment')   # 附件在字段中的位置
        self.i_verantwortlicher = code_keylist_en.index('Verantwortlicher')  # 负责人在字段中的位置
        wx.Frame.__init__(self, parent, title='Uebersicht Schadenstisch',size=(1000,500))
        self.SetBackgroundColour('LIGHT BLUE')
        # 给窗口分区，使得显示信息部分可以滚动而输入关键词的区域保持不动
        # 分区0，用于输入关键词等。固定不动
        # --------------------------------------------------------------------------------------------------------------
        # 关键词
        self.search_label = wx.StaticText(self, label='关键词：', pos=(0, 0))
        self.search_text = wx.TextCtrl(self, pos=(50, 0),style=wx.TE_PROCESS_ENTER)
        self.search_text.Bind(wx.EVT_TEXT_ENTER, lambda event, mode=1:self.search_keyword(event,mode))
        # 搜索按钮
        self.path_button = wx.Button(self, label='GO', pos=(200, 0))
        self.path_button.Bind(wx.EVT_BUTTON, lambda event, mode=1:self.search_keyword(event,mode))
        # 根据具体字段进行搜索的按钮
        self.search_detail_button = wx.Button(self, label='根据字段搜索', pos=(300, 0))
        self.search_detail_button.Bind(wx.EVT_BUTTON, self.search_detail)
        # 切换显示模式的按钮（2种显示模式，分别为只显示编号和故障信息的简略模式和显示所有内容的详细模式）
        self.switch_display_mode_button = wx.Button(self, label='显示详细内容', pos=(400, 0))
        self.switch_display_mode_button.Bind(wx.EVT_BUTTON, self.switch_display_mode)
        # 将搜索结果输出到Excel的按钮
        self.export_rot_and_gelb = wx.Button(self, label='将红/黄项目输出到Excel并发送邮件', pos=(500, 0))
        self.export_rot_and_gelb.Bind(wx.EVT_BUTTON, self.export_rot_and_gelb_to_excel)
        # 附件同步按钮
        self.file_syn_button = wx.Button(self, label='同步附件信息', pos=(800, 0))
        self.file_syn_button.Bind(wx.EVT_BUTTON, self.file_syn)

        # --------------------------------------------------------------------------------------------------------------
        # 分区1，用于显示表头。左右滑动
        # --------------------------------------------------------------------------------------------------------------
        #self.splitter_window_header = wx.SplitterWindow(self,size=(1000,50),pos=(0,50))
        #self.scrolled_window_header = wx.ScrolledWindow(self.splitter_window_header, -1, size=(1000, 50))
        self.scrolled_window_header = wx.ScrolledWindow(self, -1, size=(2500, 50), pos=(0, 50))
        self.scrolled_window_header.SetBackgroundColour(self.BackgroundColour)
        self.scrolled_window_header.SetVirtualSize((5000, 0))
        self.scrolled_window_header.SetScrollRate(20, 20)
        self.scrolled_window_header.ShowScrollbars(horz=wx.SHOW_SB_NEVER, vert=wx.SHOW_SB_NEVER)
        self.scrolled_window_header.SetBackgroundColour('LIGHT GREY')
        # --------------------------------------------------------------------------------------------------------------
        # 分区2，用于显示搜索结果。各向滑动
        # --------------------------------------------------------------------------------------------------------------
        x, y = self.GetSize()
        self.splitter_window = wx.SplitterWindow(self, size=(x, y), pos=(0, 100))
        self.scrolled_window = wx.ScrolledWindow(self.splitter_window, -1,size=(x - 18, y - 138))
        self.scrolled_window.SetBackgroundColour(self.BackgroundColour)
        self.scrolled_window.SetVirtualSize((2200, 5000))
        self.scrolled_window.SetScrollRate(20, 20)
        self.scrolled_window.Bind(wx.EVT_SCROLLWIN_THUMBRELEASE, self.scroll_syn)
        self.Bind(wx.EVT_SIZE, self.window_resize)
        # --------------------------------------------------------------------------------------------------------------
        # 绘制表格、分界线等
        # --------------------------------------------------------------------------------------------------------------
        # 显示表头的控件和分割线
        self.code_header_label = {}
        self.code_header_line = {}
        # 在显示结果的部分绘制分割线
        self.code_line_vertical = {}
        # 显示搜索结果的控件
        self.code_result_label = {}
        # 记录附件路径的控件(搜索结果)
        self.code_attachment_path = []
        # 显示结果计数
        self.keyword_found_count = 0
        # 创建控件
        self.display_mode_0()
        # --------------------------------------------------------------------------------------------------------------
        self.Centre()
        self.Show(True)
    # ------------------------------------------------------------------------------------------------------------------
    # 根据窗口大小调整显示结果区域的大小
    # ------------------------------------------------------------------------------------------------------------------
    def window_resize(self,event):
        #print(self.GetSize())
        x,y = self.GetSize()
        self.splitter_window.SetSize(x, y)
        self.scrolled_window.SetSize(x - 18, y - 138)
    # ------------------------------------------------------------------------------------------------------------------
    # 切换显示模式（简略/详细）
    # ------------------------------------------------------------------------------------------------------------------

    def switch_display_mode(self,event):
        x0, y0 = self.code_result_label[str(self.i_number) + '-' + str(0)].GetPosition()
        # 滚动视角回归原点
        self.scrolled_window_header.Scroll(x=0, y=0)
        self.scrolled_window.Scroll(x=0, y=0)
        if self.display_mode == 0:  # 简略切换成详细
            # 删除所有显示结果和横向分割线
            for n in range(self.keyword_found_count):
                if (n > 0 and self.code_result[self.i_number][n] != self.code_result[self.i_number][n - 1]) or n == 0:
                    # self.code_line_horizontal[n].Destroy()
                    for i in range(len(code_keylist_en) - 1):
                        if i == code_keylist_en.index('Problem') or i == code_keylist_en.index('Nr'):
                            self.code_result_label[str(i) + '-' + str(n)].Destroy()
                self.code_result_label[str(len(code_keylist_en) - 1) + '-' + str(n)].Destroy()
            self.display_mode = 1
            self.switch_display_mode_button.SetLabel('显示简略内容')
            for i in range(len(code_keylist_en)):
                if i == code_keylist_en.index('Problem') or i == code_keylist_en.index('Attachment') or i == code_keylist_en.index('Nr'):
                    self.code_header_label[i].Destroy()
                    self.code_line_vertical[i].Destroy()
            self.display_mode_1()
            # 滚动视角到上次浏览位置
            self.scrolled_window_header.Scroll(x=0, y=0)
            self.scrolled_window.Scroll(x=0, y=y0 / -20)
        elif self.display_mode == 1:  # 详细切换成简略
            # 删除所有显示结果和横向分割线
            for n in range(self.keyword_found_count):
                if (n > 0 and self.code_result[self.i_number][n] != self.code_result[self.i_number][n - 1]) or n == 0:
                    # self.code_line_horizontal[n].Destroy()
                    for i in range(len(code_keylist_en) - 1):
                        self.code_result_label[str(i) + '-' + str(n)].Destroy()
                self.code_result_label[str(len(code_keylist_en) - 1) + '-' + str(n)].Destroy()
            self.display_mode = 0
            self.switch_display_mode_button.SetLabel('显示详细内容')
            for i in range(len(code_keylist_en)):
                self.code_header_label[i].Destroy()
                self.code_line_vertical[i].Destroy()
            self.display_mode_0()
            # 滚动视角到上次浏览位置
            self.scrolled_window_header.Scroll(x=0, y=0)
            self.scrolled_window.Scroll(x=0, y=y0 / -20)
    def display_mode_0(self):
        for i in range(len(code_keylist_en)):
            if i == code_keylist_en.index('Problem'):
                self.code_header_label[i] = wx.StaticText(self.scrolled_window_header, label=code_keylist_cn[i],
                                                          pos=(standard_width * 1, 0),
                                                          size=(problem_width, 50))  # 部分控件的宽度比较大
                self.code_header_line[i] = wx.StaticLine(self.scrolled_window_header,
                                                         pos=(standard_width * 1 - 5, 0), size=(1, 100),
                                                         style=wx.LI_HORIZONTAL)
            elif i == code_keylist_en.index('Attachment'):
                self.code_header_label[i] = wx.StaticText(self.scrolled_window_header, label=code_keylist_cn[i],
                                                          pos=(problem_width + standard_width * 2, 0),
                                                          size=(attachment_width, 50))
                self.code_header_line[i] = wx.StaticLine(self.scrolled_window_header, pos=(
                    problem_width + standard_width * 2 - 5, 0), size=(1, 100), style=wx.LI_HORIZONTAL)
            elif i == code_keylist_en.index('Nr'):
                self.code_header_label[i] = wx.StaticText(self.scrolled_window_header, label=code_keylist_cn[i],
                                                          pos=(standard_width * 0, 0))
                self.code_header_line[i] = wx.StaticLine(self.scrolled_window_header,
                                                         pos=(standard_width * 0 - 5, 0), size=(1, 100),
                                                         style=wx.LI_HORIZONTAL)
        for i in range(len(code_keylist_en)):
            if i == code_keylist_en.index('Problem'):
                self.code_line_vertical[i] = wx.StaticLine(self.scrolled_window, pos=(standard_width * 1 - 5, 0),
                                                           size=(1, 5000), style=wx.LI_HORIZONTAL)
                # self.code_line[i].SetBackgroundColour('balck')
            elif i == code_keylist_en.index('Attachment'):
                self.code_line_vertical[i] = wx.StaticLine(self.scrolled_window, pos=(
                problem_width + standard_width * 2 - 5, 0), size=(1, 5000), style=wx.LI_HORIZONTAL)
            elif i == code_keylist_en.index('Nr'):
                self.code_line_vertical[i] = wx.StaticLine(self.scrolled_window, pos=(standard_width * 0 - 5, 0),
                                                           size=(1, 5000), style=wx.LI_HORIZONTAL)
        # 搜索结果
        for n in range(self.keyword_found_count):
            # 对于同一编号下的内容，只有第一条显示基本信息，其余只显示附件
            if (n > 0 and self.code_result[self.i_number][n] != self.code_result[self.i_number][n - 1]) or n == 0:
                for i in range(len(code_keylist_en) - 1):
                    if i == code_keylist_en.index('Problem'):
                        self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(standard_width * 1, 50 * n), size=(problem_width, 50))
                    elif i == code_keylist_en.index('Attachment'):
                        self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(problem_width + standard_width * 2, 50 * n), size=(attachment_width, 50))
                    elif i == code_keylist_en.index('Nr'):
                        self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(standard_width * 0, 50 * n))
                # 横向分割线
                self.code_line_horizontal[n] = wx.StaticLine(self.scrolled_window, pos=(0, 50 * n), size=(2500, 1), style=wx.LI_HORIZONTAL)
                    #self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(standard_width * i, 50 * n))
                # 为编号设置双击打开附件所在文件夹
                self.code_result_label[str(self.i_number) + '-' + str(n)].Bind(wx.EVT_LEFT_DCLICK, self.open_folder)
            self.code_result_label[str(self.i_attachment) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[self.i_attachment][n], pos=(problem_width + standard_width * 2, 50 * n), size=(attachment_width, 50))
            # 为附件设置双击打开附件本身
            self.code_result_label[str(self.i_attachment) + '-' + str(n)].Bind(wx.EVT_LEFT_DCLICK, self.open_file)

    def display_mode_1(self):
        for i in range(len(code_keylist_en)):
            if i == code_keylist_en.index('Problem'):
                self.code_header_label[i] = wx.StaticText(self.scrolled_window_header, label=code_keylist_cn[i],
                                                          pos=(standard_width * i, 0),
                                                          size=(problem_width, 50))  # 部分控件的宽度比较大
                self.code_header_line[i] = wx.StaticLine(self.scrolled_window_header,
                                                         pos=(standard_width * i - 5, 0), size=(1, 100),
                                                         style=wx.LI_HORIZONTAL)
                # self.code_header_line[i].SetBackgroundColour('balck')
            elif i > code_keylist_en.index('Problem') and i < code_keylist_en.index('Bemerkung'):
                self.code_header_label[i] = wx.StaticText(self.scrolled_window_header, label=code_keylist_cn[i],
                                                          pos=(problem_width + standard_width * i, 0))
                self.code_header_line[i] = wx.StaticLine(self.scrolled_window_header,
                                                         pos=(problem_width + standard_width * i - 5, 0),
                                                         size=(1, 100), style=wx.LI_HORIZONTAL)
            elif i == code_keylist_en.index('Bemerkung'):
                self.code_header_label[i] = wx.StaticText(self.scrolled_window_header, label=code_keylist_cn[i],
                                                          pos=(problem_width + standard_width * i, 0),
                                                          size=(bemerkung_width, 50))
                self.code_header_line[i] = wx.StaticLine(self.scrolled_window_header,
                                                         pos=(problem_width + standard_width * i - 5, 0),
                                                         size=(1, 100), style=wx.LI_HORIZONTAL)
            elif i > code_keylist_en.index('Bemerkung') and i < code_keylist_en.index('Attachment'):
                self.code_header_label[i] = wx.StaticText(self.scrolled_window_header, label=code_keylist_cn[i],
                                                          pos=(
                                                              bemerkung_width + problem_width + standard_width * i, 0))
                self.code_header_line[i] = wx.StaticLine(self.scrolled_window_header, pos=(
                    bemerkung_width + problem_width + standard_width * i - 5, 0), size=(1, 100), style=wx.LI_HORIZONTAL)
            elif i == code_keylist_en.index('Attachment'):
                self.code_header_label[i] = wx.StaticText(self.scrolled_window_header, label=code_keylist_cn[i],
                                                          pos=(problem_width + bemerkung_width + standard_width * i, 0),
                                                          size=(attachment_width, 50))
                self.code_header_line[i] = wx.StaticLine(self.scrolled_window_header, pos=(
                    bemerkung_width + problem_width + standard_width * i - 5, 0), size=(1, 100), style=wx.LI_HORIZONTAL)
            else:
                self.code_header_label[i] = wx.StaticText(self.scrolled_window_header, label=code_keylist_cn[i],
                                                          pos=(standard_width * i, 0))
                self.code_header_line[i] = wx.StaticLine(self.scrolled_window_header,
                                                         pos=(standard_width * i - 5, 0), size=(1, 100),
                                                         style=wx.LI_HORIZONTAL)
        for i in range(len(code_keylist_en)):
            if i == code_keylist_en.index('Problem'):
                self.code_line_vertical[i] = wx.StaticLine(self.scrolled_window, pos=(standard_width * i - 5, 0),
                                                           size=(1, 5000), style=wx.LI_HORIZONTAL)
                # self.code_line[i].SetBackgroundColour('balck')
            elif i > code_keylist_en.index('Problem') and i < code_keylist_en.index('Bemerkung'):
                self.code_line_vertical[i] = wx.StaticLine(self.scrolled_window,
                                                           pos=(problem_width + standard_width * i - 5, 0),
                                                           size=(1, 5000), style=wx.LI_HORIZONTAL)
            elif i == code_keylist_en.index('Bemerkung'):
                self.code_line_vertical[i] = wx.StaticLine(self.scrolled_window,
                                                           pos=(problem_width + standard_width * i - 5, 0),
                                                           size=(1, 5000), style=wx.LI_HORIZONTAL)
            elif i > code_keylist_en.index('Bemerkung') and i < code_keylist_en.index('Attachment'):
                self.code_line_vertical[i] = wx.StaticLine(self.scrolled_window, pos=(
                bemerkung_width + problem_width + standard_width * i - 5, 0), size=(1, 5000),
                                                           style=wx.LI_HORIZONTAL)
            elif i == code_keylist_en.index('Attachment'):
                self.code_line_vertical[i] = wx.StaticLine(self.scrolled_window, pos=(
                bemerkung_width + problem_width + standard_width * i - 5, 0), size=(1, 5000),
                                                           style=wx.LI_HORIZONTAL)
            else:
                self.code_line_vertical[i] = wx.StaticLine(self.scrolled_window, pos=(standard_width * i - 5, 0),
                                                           size=(1, 5000), style=wx.LI_HORIZONTAL)
        # 搜索结果
        for n in range(self.keyword_found_count):
            # 对于同一编号下的内容，只有第一条显示基本信息，其余只显示附件
            if (n > 0 and self.code_result[self.i_number][n] != self.code_result[self.i_number][n - 1]) or n == 0:
                for i in range(len(code_keylist_en) - 1):
                    if i == code_keylist_en.index('Problem'):
                        self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(standard_width * i, 50 * n), size=(problem_width, 50))
                    elif i > code_keylist_en.index('Problem') and i < code_keylist_en.index('Bemerkung'):
                        self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(problem_width + standard_width * i, 50 * n))
                    elif i == code_keylist_en.index('Bemerkung'):
                        self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(problem_width + standard_width * i, 50 * n), size=(bemerkung_width, 50))
                    elif i > code_keylist_en.index('Bemerkung') and i < code_keylist_en.index('Attachment'):
                        self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(bemerkung_width + problem_width + standard_width * i, 50 * n))
                    elif i == code_keylist_en.index('Attachment'):
                        self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(problem_width + bemerkung_width + standard_width * i, 50 * n), size=(attachment_width, 50))
                    else:
                        self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(standard_width * i, 50 * n))
                # 分割线
                self.code_line_horizontal[n] = wx.StaticLine(self.scrolled_window, pos=(0, 50 * n), size=(2500, 1), style=wx.LI_HORIZONTAL)
                    #self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(standard_width * i, 50 * n))
                # 为编号设置双击打开附件所在文件夹
                self.code_result_label[str(self.i_number) + '-' + str(n)].Bind(wx.EVT_LEFT_DCLICK, self.open_folder)
            self.code_result_label[str(self.i_attachment) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[self.i_attachment][n], pos=(problem_width + bemerkung_width + standard_width * self.i_attachment, 50 * n), size=(attachment_width, 50))
            # 为附件设置双击打开附件本身
            self.code_result_label[str(self.i_attachment) + '-' + str(n)].Bind(wx.EVT_LEFT_DCLICK, self.open_file)

    # ------------------------------------------------------------------------------------------------------------------
    # 打开根据具体字段搜索的字窗口
    # ------------------------------------------------------------------------------------------------------------------
    def search_detail(self,event):
        print('open sub window')
        frame = SubWindow(None, 'Small editor')
        frame.Show(True)
    # ------------------------------------------------------------------------------------------------------------------
    # 同步附件
    # ------------------------------------------------------------------------------------------------------------------
    def file_syn(self,event):
        file2txt()
    # ------------------------------------------------------------------------------------------------------------------
    # 滚动条同步,使得分别位于2个分区的表头和搜索结果能够在横向同步滚动
    # ------------------------------------------------------------------------------------------------------------------
    def scroll_syn(self,event):
        x0,y0 = self.code_result_label[str(self.i_number) + '-' + str(0)].GetPosition()
        #print(self.code_number_label[0].GetPosition())
        self.scrolled_window_header.Scroll(x=x0 / -20,y=0)
    # ------------------------------------------------------------------------------------------------------------------
    # 在进行搜索前,将之前的搜索结果清空
    # ------------------------------------------------------------------------------------------------------------------
    def search_reset(self):
        global code_keylist_en, code_value, keyword_count, code_keyword
        # 滚动视角回归原点
        self.scrolled_window_header.Scroll(x=0, y=0)
        self.scrolled_window.Scroll(x=0, y=0)
        #print(self.keyword_found_count)
        # 删除所有显示结果和横向分割线
        for n in range(self.keyword_found_count):
            if (n > 0 and self.code_result[self.i_number][n] != self.code_result[self.i_number][n - 1]) or n == 0:
                #self.code_line_horizontal[n].Destroy()
                for i in range(len(code_keylist_en) - 1):
                    self.code_result_label[str(i) + '-' + str(n)].Destroy()
            self.code_result_label[str(len(code_keylist_en) - 1) + '-' + str(n)].Destroy()
        # 清空变量里的数据
        self.code_result_label = {}
        self.code_result = [[] for i in range(len(code_keylist_en))]
        self.keyword_found_n_list = []
        self.code_attachment_path = []
        self.code_line_horizontal = {}
        # 显示模式回归简略模式
        self.display_mode = 0
    # ------------------------------------------------------------------------------------------------------------------
    def search_keyword(self,event,mode):
        global code_keylist_en, code_value, keyword_count, code_keyword, keyword
        # 如果之前有过搜索记录，则界面上的删除记录，并将视角重置
        self.search_reset()
        #  搜索信息
        # --------------------------------------------------------------------------------------------------------------
        print('searching with mode ' + str(mode))
        if mode == 1:   # 模糊搜索，即输入一个关键词，对所有字段都进行该关键词的搜索
            # 获取关键词
            if keyword == '':
                keyword = self.search_text.GetValue().lower()
            # 遍历各个附件，对其基本信息和文件内容进行搜索
            for n in range(len(file_path_list)):
                # 如果该附件所对应的编号已在结果列表中，则此条记入搜索结果中
                if code_value[self.i_number][n] in self.code_result[self.i_number]:
                    self.keyword_found(n=n)
                    continue
                # 对基本信息进行搜索
                for i in range(len(code_keylist_en)):
                    if keyword in str(code_value[i][n]).lower() and not (n in self.keyword_found_n_list):
                        self.keyword_found(n=n)
                        break
                # 搜索文档内容
                txt_path = txt_folder + file_folder_list[n] + '\\' + file_name_list[n] + '.txt'  # txt路径
                if os.path.exists(txt_path):  # 如果有该txt文件，即该文档支持搜索
                    if file_extension_list[n] == 'docx':
                        txt_file = open(txt_path)
                    else:
                        txt_file = open(txt_path, encoding='utf-8')
                    # print(txt_path)
                    # 遍历文件内容
                    for i_line, line in enumerate(txt_file.readlines()):
                        if keyword in line.lower():
                            if not (n in self.keyword_found_n_list):
                                self.keyword_found(n=n)
                                break
            keyword = ''
        elif mode == 2: # 根据字段检索，只有所有关键词都符合才符合（与）
            print(code_keyword)
            # 遍历各个附件，对其基本信息和文件内容进行搜索
            for n in range(len(file_path_list)):
                # 如果该附件所对应的编号已在结果列表中，则此条记入搜索结果中
                if code_value[self.i_number][n] in self.code_result[self.i_number]:
                    self.keyword_found(n=n)
                    continue
                # 对基本信息进行搜索，但不包括附件名称（等搜索文档内容时再对名称进行检查）
                keyword_found_count_temp = 0    # 用于统计多少字段的关键词符合
                for i in range(len(code_keylist_en) - 1):
                    if code_keyword[i] in str(code_value[i][n]).lower() and code_keyword[i] != '':
                        keyword_found_count_temp += 1
                # 搜索文档内容
                if code_keyword[self.i_attachment] != '':
                    # 检查文件名
                    if code_keyword[self.i_attachment] in file_name_list[n].lower() and not (n in self.keyword_found_n_list):
                        keyword_found_count_temp += 1
                    else:
                        # 检查文件内容
                        txt_path = txt_folder + file_folder_list[n] + '\\' + file_name_list[n] + '.txt'  # txt路径
                        if os.path.exists(txt_path):  # 如果有该txt文件，即该文档支持搜索
                            if file_extension_list[n] == 'docx':
                                txt_file = open(txt_path)
                            else:
                                txt_file = open(txt_path, encoding='utf-8')
                            # print(txt_path)
                            # 遍历文件内容
                            for i_line, line in enumerate(txt_file.readlines()):
                                if code_keyword[self.i_attachment] in line.lower():
                                    if not (n in self.keyword_found_n_list):
                                        keyword_found_count_temp += 1
                                        break
                # 统计多少字段的关键词是符合的。与输入的关键词数量一致，即全部满足才算符合
                if keyword_found_count_temp == keyword_count:
                    self.keyword_found(n=n)

        elif mode == 3: # 根据字段检索，只要有其中一个关键词符合就可以（或）
            print(code_keyword)
            # 遍历各个附件，对其基本信息和文件内容进行搜索
            for n in range(len(file_path_list)):
                # 如果该附件所对应的编号已在结果列表中，则此条记入搜索结果中
                if code_value[self.i_number][n] in self.code_result[self.i_number]:
                    self.keyword_found(n=n)
                    continue
                # 对基本信息进行搜索
                for i in range(len(code_keylist_en)):
                    if code_keyword[i] in str(code_value[i][n]).lower() and not (n in self.keyword_found_n_list) and code_keyword[i] != '':
                        self.keyword_found(n=n)
                        break
                # 搜索文档内容
                if code_keyword[self.i_attachment] != '':
                    txt_path = txt_folder + file_folder_list[n] + '\\' + file_name_list[n] + '.txt'  # txt路径
                    # 检查文件内容
                    if os.path.exists(txt_path):  # 如果有该txt文件，即该文档支持搜索
                        if file_extension_list[n] == 'docx':
                            txt_file = open(txt_path)
                        else:
                            txt_file = open(txt_path, encoding='utf-8')
                        # print(txt_path)
                        # 遍历文件内容
                        for i_line, line in enumerate(txt_file.readlines()):
                            if code_keyword[self.i_attachment] in line.lower():
                                if not (n in self.keyword_found_n_list):
                                    self.keyword_found(n=n)
                                    break

        # --------------------------------------------------------------------------------------------------------------
        # 显示搜索结果
        # --------------------------------------------------------------------------------------------------------------
        # 统计结果个数
        self.keyword_found_count = len(self.keyword_found_n_list)
        print(self.keyword_found_n_list)
        #print(self.code_attachment)
        # 搜索结果
        if self.display_mode == 0:
            for n in range(self.keyword_found_count):
                # 对于同一编号下的内容，只有第一条显示基本信息，其余只显示附件
                if (n > 0 and self.code_result[self.i_number][n] != self.code_result[self.i_number][n - 1]) or n == 0:
                    for i in range(len(code_keylist_en) - 1):
                        if i == code_keylist_en.index('Problem'):
                            self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(standard_width * 1, 50 * n), size=(problem_width, 50))
                        elif i == code_keylist_en.index('Attachment'):
                            self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(problem_width + standard_width * 2, 50 * n), size=(attachment_width, 50))
                        elif i == code_keylist_en.index('Nr'):
                            self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(standard_width * 0, 50 * n))
                    # 横向分割线
                    self.code_line_horizontal[n] = wx.StaticLine(self.scrolled_window, pos=(0, 50 * n), size=(2500, 1), style=wx.LI_HORIZONTAL)
                        #self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(standard_width * i, 50 * n))
                    # 为编号设置双击打开附件所在文件夹
                    self.code_result_label[str(self.i_number) + '-' + str(n)].Bind(wx.EVT_LEFT_DCLICK, self.open_folder)
                self.code_result_label[str(self.i_attachment) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[self.i_attachment][n], pos=(problem_width + standard_width * 2, 50 * n), size=(attachment_width, 50))
                # 为附件设置双击打开附件本身
                self.code_result_label[str(self.i_attachment) + '-' + str(n)].Bind(wx.EVT_LEFT_DCLICK, self.open_file)
        elif self.display_mode == 1:
            for n in range(self.keyword_found_count):
                # 对于同一编号下的内容，只有第一条显示基本信息，其余只显示附件
                if (n > 0 and self.code_result[self.i_number][n] != self.code_result[self.i_number][n - 1]) or n == 0:
                    for i in range(len(code_keylist_en) - 1):
                        if i == code_keylist_en.index('Problem'):
                            self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(standard_width * i, 50 * n), size=(problem_width, 50))
                        elif i > code_keylist_en.index('Problem') and i < code_keylist_en.index('Bemerkung'):
                            self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(problem_width + standard_width * i, 50 * n))
                        elif i == code_keylist_en.index('Bemerkung'):
                            self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(problem_width + standard_width * i, 50 * n), size=(bemerkung_width, 50))
                        elif i > code_keylist_en.index('Bemerkung') and i < code_keylist_en.index('Attachment'):
                            self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(bemerkung_width + problem_width + standard_width * i, 50 * n))
                        elif i == code_keylist_en.index('Attachment'):
                            self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(problem_width + bemerkung_width + standard_width * i, 50 * n), size=(attachment_width, 50))
                        else:
                            self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(standard_width * i, 50 * n))
                    # 分割线
                    self.code_line_horizontal[n] = wx.StaticLine(self.scrolled_window, pos=(0, 50 * n), size=(2500, 1), style=wx.LI_HORIZONTAL)
                        #self.code_result_label[str(i) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[i][n], pos=(standard_width * i, 50 * n))
                    # 为编号设置双击打开附件所在文件夹
                    self.code_result_label[str(self.i_number) + '-' + str(n)].Bind(wx.EVT_LEFT_DCLICK, self.open_folder)
                self.code_result_label[str(self.i_attachment) + '-' + str(n)] = wx.StaticText(self.scrolled_window, label=self.code_result[self.i_attachment][n], pos=(problem_width + bemerkung_width + standard_width * self.i_attachment, 50 * n), size=(attachment_width, 50))
            # 为附件设置双击打开附件本身
            self.code_result_label[str(self.i_attachment) + '-' + str(n)].Bind(wx.EVT_LEFT_DCLICK, self.open_file)
        # 重新设置滚动窗口的范围
        self.scrolled_window.SetVirtualSize((2200, self.keyword_found_count * 50))
        # --------------------------------------------------------------------------------------------------------------
    # ------------------------------------------------------------------------------------------------------------------
    # 函数。作用是将发现有关键词的条目的内容存放到变量中。
    # ------------------------------------------------------------------------------------------------------------------
    def keyword_found(self,n):
        global code_keylist_en, code_value, keyword_count, code_keyword
        self.keyword_found_n_list.append(n)
        for i in range(len(code_keylist_en)):
            self.code_result[i].append(code_value[i][n])
        self.code_attachment_path.append(file_path_list[n])

    # ------------------------------------------------------------------------------------------------------------------
    # 打开附件所在文件夹
    # ------------------------------------------------------------------------------------------------------------------
    def open_folder(self, event):
        x0, y0 = self.GetPosition()
        x, y = wx.GetMousePosition()
        y1 = self.scrolled_window.CalcUnscrolledPosition(x, y)[1]
        print(y0, y, y1)
        i_code_selected = round((y1 - 150 - y0) / 50)
        print(i_code_selected, self.code_result[self.i_number][i_code_selected])
        folder_path = storage_root_path + '\\' + self.code_result[self.i_number][i_code_selected] + '\\'
        os.startfile(folder_path)
    # ------------------------------------------------------------------------------------------------------------------
    # 打开附件
    # ------------------------------------------------------------------------------------------------------------------
    def open_file(self, event):
        x0, y0 = self.GetPosition()
        x, y = wx.GetMousePosition()
        y1 = self.scrolled_window.CalcUnscrolledPosition(x, y)[1]
        #print(y0, y, y1)
        i_code_selected = round((y1 - 150 - y0) / 50)
        print(i_code_selected, self.code_result[self.i_attachment][i_code_selected])
        file_path = self.code_attachment_path[i_code_selected]
        os.startfile(file_path)
    # ------------------------------------------------------------------------------------------------------------------
    # 输出当前结果到Excel并发送邮件
    # ------------------------------------------------------------------------------------------------------------------
    def export_current_to_excel(self, event):
        wbk = openpyxl.load_workbook('Uebersicht Schadenstisch Template.xlsx')
        sht = wbk.active
        x = 0
        for n in range(self.keyword_found_count):
            # 对于同一编号下的内容，只有第一条显示基本信息，其余只显示附件
            if (n > 0 and self.code_result[self.i_number][n] != self.code_result[self.i_number][n - 1]) or n == 0:
                for i in range(len(code_keylist_en) - 1):
                    sht.cell(row=3 + x, column=1 + i).value = self.code_result[i][n]
                x += 1
        result_path = os.getcwd() + '/03_Result' + '/Uebersicht Schadenstisch.xlsx'
        wbk.save(result_path)
        outlook = win32com.client.Dispatch('outlook.application')
        receivers = ['han_hou333@163.com']
        contact_name_list = []
        contact_mail_list = []
        mail = outlook.CreateItem(0)
        mail.To = receivers[0]
        # mail.Recipients.Add(receivers[1])
        mail.Subject = 'this is a test'
        mail.Body = 'test test test'
        # 添加附件
        mail.Attachments.Add(result_path)
        mail.Send()
        dlg = wx.MessageDialog(self, '发送完毕', 'Info', style=wx.OK)
        if dlg.ShowModal() == wx.ID_OK:
            dlg.Destroy()
    # ------------------------------------------------------------------------------------------------------------------
    # 输出所有红色和黄色状态的结果到Excel并发送邮件
    # ------------------------------------------------------------------------------------------------------------------
    def export_rot_and_gelb_to_excel(self, event):
        wbk = openpyxl.load_workbook('Uebersicht Schadenstisch Template.xlsx')
        sht = wbk.active
        receivers = ''
        x = 0
        for n, item in enumerate(status_list):
            if item == 'gelb' or item == 'rot':
                # 录入信息
                for i in range(len(code_keylist_en) - 1):
                    sht.cell(row=3 + x, column=1 + i).value = code_value_sht[i][n]
                x += 1
                # 添加收件人名单
                contact_name = verantwortlicher_list[n]
                contact_mail = contact_mail_list[contact_name_list.index(contact_name)]
                if contact_mail not in receivers:
                    receivers = receivers + ';' + contact_mail
        receivers = receivers[1:]
        result_path = os.getcwd() + '/03_Result' + '/Uebersicht Schadenstisch.xlsx'
        wbk.save(result_path)
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        #print(receivers)
        mail.To = receivers
        # mail.Recipients.Add(receivers[1])
        mail.Subject = 'this is a test'
        mail.Body = 'test test test'
        # 添加附件
        mail.Attachments.Add(result_path)
        mail.Send()
        dlg = wx.MessageDialog(self, '发送完毕', 'Info', style=wx.OK)
        if dlg.ShowModal() == wx.ID_OK:
            dlg.Destroy()
    # ------------------------------------------------------------------------------------------------------------------

# 子窗口，用于输入根据字段搜索的关键词
class SubWindow(wx.Frame):
    # 窗口初始化
    # ------------------------------------------------------------------------------------------------------------------
    def __init__(self,parent,title):
        wx.Frame.__init__(self, parent, title='Uebersicht Schadenstisch',size=(500,1000))
        # 给窗口分区，使得显示信息部分可以滚动而输入关键词的区域保持不动
        # 分区0，用于输入关键词等。固定不动
        # ------------------------------------------------------------------------------------------------------------------
        #self.search_text.Bind(wx.EVT_TEXT_ENTER, self.search_keyword)
        # 搜索按钮
        self.path_button = wx.Button(self, label='GO', pos=(350, 0))
        self.path_button.Bind(wx.EVT_BUTTON, self.search_keyword)
        # 模式选择
        self.mode_selection = wx.RadioBox(self, label='搜索模式', pos=(350, 100), choices=['与', '或'])
        #self.file_syn_button.Bind(wx.EVT_BUTTON, self.file_syn)
        # 分区1，用于显示表头。左右滑动
        # --------------------------------------------------------------------------------------------------------------
        # 显示表头的控件
        standard_gap = 50
        self.code_header_label = {}
        self.code_keyword_text = {}
        # 初始化控件
        for i in range(len(code_keylist_en)):
            self.code_header_label[i] = wx.StaticText(self, label=code_keylist_en[i], pos=(0, standard_gap * i))
            self.code_keyword_text[i] = wx.TextCtrl(self, pos=(100, standard_gap * i), size=(200,30))
            self.code_keyword_text[i].Bind(wx.EVT_TEXT_ENTER, self.search_keyword)

        self.Bind(wx.EVT_CLOSE, self.show_main_window)
        self.Centre()
    def show_main_window(self, event):
        self.Destroy()
        frame_main_window.Show(True)
    # 获取输入的关键词，并计算关键词的个数
    def search_keyword(self,event):
        global keyword_count, code_keyword
        keyword_count = 0
        for i in range(len(code_keylist_en)):
            code_keyword[i] = str(self.code_keyword_text[i].GetLineText(0)).lower()
            if code_keyword[i] != '':
                keyword_count += 1
        self.Destroy()
        # 调用主窗口搜索关键词的函数，并传递参数
        #frame_main_window = MainWindow(None, 'Small editor')
        frame_main_window.Show(True)
        frame_main_window.search_keyword(event=1,mode=self.mode_selection.GetSelection() + 2)
        #aaa = MainWindow.search_keyword(self=MainWindow,event,mode=1)


class PreWindow(wx.Frame):
    # 先遍历所有文件并将文件路径存放在变量中
    check_files(storage_root_path)
    # 窗口初始化
    # ------------------------------------------------------------------------------------------------------------------
    def __init__(self,parent,title):
        wx.Frame.__init__(self, parent, title='Uebersicht Schadenstisch',size=(1000, 500))
        # 关键词
        #self.search_label = wx.StaticText(self, label='关键词：', pos=(300, 0))
        self.search_text = wx.TextCtrl(self, pos=(200, 300),size=(400, 50), style=wx.TE_PROCESS_ENTER)
        self.search_text.Bind(wx.EVT_TEXT_ENTER, self.search_keyword)
        # 搜索按钮
        self.path_button = wx.Button(self, label='GO', pos=(650, 300),size=(200, 50))
        self.path_button.Bind(wx.EVT_BUTTON, self.search_keyword)
        # 根据具体字段进行搜索的按钮
        self.search_detail_button = wx.Button(self, label='根据字段搜索', pos=(650, 380),size=(200, 30))
        self.search_detail_button.Bind(wx.EVT_BUTTON, self.search_detail)
        # 图片
        pic_svw_path = ui_material_folder + '\\SVW.jpg'
        pic_svw_image = wx.Image(pic_svw_path, wx.BITMAP_TYPE_JPEG)
        pic_svw_temp = pic_svw_image.ConvertToBitmap()
        pic_svw_bitmap = wx.StaticBitmap(self, bitmap=pic_svw_temp, pos=(0, 0), size=(500, 200))

        # --------------------------------------------------------------------------------------------------------------
        self.Centre()
        self.Show(True)
        self.Bind(wx.EVT_CLOSE, self.show_main_window)
    def show_main_window(self,event):
        self.Destroy()
        frame_main_window.Destroy()

    # ------------------------------------------------------------------------------------------------------------------
    # # 调用主窗口搜索关键词的函数，并传递参数
    # ------------------------------------------------------------------------------------------------------------------
    def search_keyword(self, event):
        global keyword
        self.Destroy()
        print('open sub window')
        keyword = self.search_text.GetValue().lower()
        #frame_main_window = MainWindow(None, 'Small editor')
        frame_main_window.Show(True)
        frame_main_window.search_keyword(event=1, mode=1)
    # ------------------------------------------------------------------------------------------------------------------
    # 打开根据具体字段搜索的字窗口
    # ------------------------------------------------------------------------------------------------------------------
    def search_detail(self, event):
        self.Destroy()
        print('open sub window')
        frame_sub_window = SubWindow(None, 'Small editor')
        frame_sub_window.Show(True)
# ----------------------------------------------------------------------------------------------------------------------
# 进入消息循环
app = wx.App()
frame_main_window = MainWindow(None, 'Small editor')
frame_main_window.Show(False)
frame_pre_window = PreWindow(None, 'Small editor')
app.MainLoop()


class MainWindow(wx.Frame):
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, title='读取MDF文件并将数据录入至Access', size=(500, 160))
        # 选择文件的按钮
        self.path_button = wx.Button(self, label='选择文件', pos=(0, 0))
        self.path_button.Bind(wx.EVT_BUTTON, self.get_file_path)
        # 状态栏
        self.status_label = wx.StaticText(self, label='', pos=(200, 40))
        # 进度条
        self.progress_label = wx.StaticText(self, label='', pos=(250, 70))
        self.progress_bar = wx.Gauge(self, 1001, 100, pos=(0, 100), size=(500, 20), style=wx.GA_HORIZONTAL)
        self.progress_bar.Show(False)
        # 显示介面并打开选择文件的对话框
        self.Show(True)
        self.get_file_path(event='')

    def get_file_path(self, event):
        dlg = wx.FileDialog(self ,message='选择MDF文件', defaultDir=os.getcwd(), wildcard='MDF files (.MDF)|*.MDF')
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            self.update_data(path)
        dlg.Destroy()