import templates
import pandas as pd


class FormatData():

    def __init__(self, filename):
        self.filename = filename  
        self.df = pd.read_excel(self.filename)
        self.names = self.df['Имя'].tolist()
        self.questions = self.df['Вопрос'].tolist()
        self.datapaths = self.df['Путь к данным'].tolist()
        self.codes = self.df['Выражение'].tolist()
        self.example_table = {ord('"') : None, ord("'") : None, ord(' ') : None}
        self.example_code_table = {ord('&') : '&amp;', 
        ord('<') : '&lt;', ord('>') : '&gt;' }

    def make_sql_request(self):
        '''Возвращает строку данных сформированного sql запрорса
        с данными из таблицы для подстановки в общий шаблон'''
        l = []        
        for name in self.names:
            l.append(templates.request.replace('?!имя', name.translate(self.example_table)))
        self.sql_request = ',\n'.join(l)
        return self.sql_request

    def format_code(self):
        '''Форматирует выражения в списке для формата xml'''
        self.format_codes = []
        for code in self.codes:
            string_code = str(code)
            self.format_codes.append(string_code.translate(self.example_code_table))
        return self.format_codes 

    def make_calculated(self):
        '''Возвращает шаблон калькулируемых сгруппированых калькулируемых
        данных для подстановки в общий шаблон'''
        replaced_list = []
        for code, datapath in zip(self.format_codes, self.datapaths):
            replaced_list.append(templates.calculated.replace('?!код', code).replace(
                '?!заголовок', datapath).replace('?!путь', datapath.translate(self.example_table)))
        self.calculated = '\n'.join(replaced_list)
        return self.calculated    

    def replace_data(self, template):
        '''Возвращает сформированные шаблоны с данными из таблицы'''
        replaced_list = []
        for name, question in zip(self.names, self.questions):
            replaced_list.append(template.replace(
                '?!имя', name).replace('?!вопрос', question))  
        self.replaced_data = '\n'.join(replaced_list)    
        return self.replaced_data

    def make_full_data(self):
        sql = self.make_sql_request()
        self.format_code()
        calculated_data = self.make_calculated()
        dataset = self.replace_data(templates.dataset)
        parameter = self.replace_data(templates.parameter)
        selected = self.replace_data(templates.selected)
        settings = self.replace_data(templates.settings)
        full_data = templates.root.replace('<!--ВставкаШаблонаНабораДанных-->', 
        dataset).replace('<!--ВставкаЗапроса-->', sql
        ).replace('<!--ВставкаШаблонаВычесляемыхВыражений-->', calculated_data
        ).replace('<!--ВставкаШаблонаПараметров-->', parameter
        ).replace('<!--ВставкаШаблонаВыборки-->', selected
        ).replace('<!--ВставкаШаблонаНастроек-->', settings)
        return full_data
     