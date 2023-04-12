import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
from openpyxl import Workbook
import xlsxwriter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
import os

# Считать файл
df = pd.read_excel("данные.xlsx", sheet_name='Данные')

# Подключение списка признаков и ИБ
tag = set(df.Признак.drop_duplicates())
MH = set(df.ИБ.drop_duplicates())


def do_DB_Alternatives(result):
    global exemplares_1, exemplares_2, exemplares_3, exemplares_4, exemplares_5

    # Перебираем истории болезни
    alter_name1 = 1
    alter_name2 = 1
    alter_name3 = 1
    alter_name4 = 1
    alter_name5 = 1

    exemplares_1 = []
    exemplares_2 = []
    exemplares_3 = []
    exemplares_4 = []
    exemplares_5 = []

    for mh in MH:
        result = df.loc[df['ИБ'] == mh]
        diagnos = result.iloc[0]['Заболевание']
        for tags in tag:

            # 1 and 2 periods
            result1 = result.loc[result['Признак'] == tags]
            x = list(result1.Момент)
            y = list(result1.Значение)

            for i in x:
                # Делим моменты
                index = x.index(i)
                x_list_1 = x[:index + 1]
                x_list_2 = x[index + 1:]
                y_list_1 = y[:index + 1]
                y_list_2 = y[index + 1:]

                # Проверяем правильность разделения
                intersection = set(y_list_1) & set(y_list_2)
                if not intersection:
                    if y_list_2 and len(x_list_1) < 4 and len(x_list_2) < 4:
                        # 2 PERIODS
                        # Запись в датафрейм
                        exemplares_2.append(
                            {'diagnos': diagnos, 'name': f'альтернатива_2.{str(alter_name2)}', 'MH': f'{mh}',
                             'tag': f'{tags}', 'NPD': 2,
                             'Values_PD_1': set(y_list_1), 'lower_bound_1': int(x_list_1[0]),
                             'upper_bound_1': int(x_list_1[-1]), 'Values_PD_2': set(y_list_2),
                             'lower_bound_2': int(int(x_list_2[0]) - int(x_list_1[-1])),
                             'upper_bound_2': int(int(x_list_2[-1]) - int(x_list_1[-1]))
                             })
                        alter_name2 = alter_name2 + 1
                    elif not y_list_2:
                        # 1 PERIOD
                        # Запись в датафрейм
                        exemplares_1.append(
                            {'diagnos': diagnos, 'name': f'альтернатива_1.{str(alter_name1)}', 'MH': f'{mh}',
                             'tag': f'{tags}', 'NPD': 1,
                             'Values_PD': set(y_list_1), 'lower_bound': int(x_list_1[0]),
                             'upper_bound': int(x_list_1[-1])})
                        alter_name1 = alter_name1 + 1

            # 3 PERIODS

            for i in x:
                # Делим моменты
                index = x.index(i)
                for j in x[1:]:
                    index2 = x.index(j)
                    x_list_1 = x[:index + 1]
                    x_list_2 = x[index + 1:index2 + 1]
                    x_list_3 = x[index2 + 1:]
                    y_list_1 = y[:index + 1]
                    y_list_2 = y[index + 1:index2 + 1]
                    y_list_3 = y[index2 + 1:]

                    # Проверяем правильность разделения
                    intersection = set(y_list_1) & set(y_list_2)
                    intersection2 = set(y_list_2) & set(y_list_3)

                    # Запись в датафрейм
                    if not intersection and not intersection2 and y_list_1 and y_list_2 and y_list_3 and len(
                            x_list_1) < 4 and len(x_list_2) < 4 and len(x_list_3) < 4:
                        exemplares_3.append(
                            {'diagnos': diagnos, 'name': f'альтернатива_3.{str(alter_name3)}', 'MH': f'{mh}',
                             'tag': f'{tags}', 'NPD': 3,
                             'Values_PD_1': set(y_list_1), 'lower_bound_1': int(x_list_1[0]),
                             'upper_bound_1': int(x_list_1[-1]), 'Values_PD_2': set(y_list_2),
                             'lower_bound_2': int(int(x_list_2[0]) - int(x_list_1[-1])),
                             'upper_bound_2': int(int(x_list_2[-1]) - int(x_list_1[-1])),
                             'Values_PD_3': set(y_list_3),
                             'lower_bound_3': int(int(x_list_3[0]) - int(x_list_2[-1])),
                             'upper_bound_3': int(int(x_list_3[-1]) - int(x_list_2[-1]))})
                        alter_name3 = alter_name3 + 1

            # 4 PERIODS

            for i in x:
                # Делим моменты
                index = x.index(i)
                for j in x[1:]:
                    index2 = x.index(j)
                    for k in x[2:]:
                        index3 = x.index(k)
                        x_list_1 = x[:index + 1]
                        x_list_2 = x[index + 1:index2 + 1]
                        x_list_3 = x[index2 + 1:index3 + 1]
                        x_list_4 = x[index3 + 1:]

                        y_list_1 = y[:index + 1]
                        y_list_2 = y[index + 1:index2 + 1]
                        y_list_3 = y[index2 + 1:index3 + 1]
                        y_list_4 = y[index3 + 1:]

                        # Проверяем правильность разделения
                        intersection = set(y_list_1) & set(y_list_2)
                        intersection2 = set(y_list_2) & set(y_list_3)
                        intersection3 = set(y_list_3) & set(y_list_4)

                        # Запись в датафрейм
                        if not intersection and not intersection2 and not intersection3 and y_list_1 and y_list_2 and y_list_3 and y_list_4 and len(
                                x_list_1) < 4 and len(x_list_2) < 4 and len(x_list_3) < 4 and len(x_list_4) < 4:
                            exemplares_4.append(
                                {'diagnos': diagnos, 'name': f'альтернатива_4.{str(alter_name4)}', 'MH': f'{mh}',
                                 'tag': f'{tags}',
                                 'NPD': 4,
                                 'Values_PD_1': set(y_list_1), 'lower_bound_1': int(x_list_1[0]),
                                 'upper_bound_1': int(x_list_1[-1]), 'Values_PD_2': set(y_list_2),
                                 'lower_bound_2': int(int(x_list_2[0]) - int(x_list_1[-1])),
                                 'upper_bound_2': int(int(x_list_2[-1]) - int(x_list_1[-1])),
                                 'Values_PD_3': set(y_list_3),
                                 'lower_bound_3': int(int(x_list_3[0]) - int(x_list_2[-1])),
                                 'upper_bound_3': int(int(x_list_3[-1]) - int(x_list_2[-1])),
                                 'Values_PD_4': set(y_list_4),
                                 'lower_bound_4': int(int(x_list_4[0]) - int(x_list_3[-1])),
                                 'upper_bound_4': int(int(x_list_4[-1]) - int(x_list_3[-1]))})
                            alter_name4 = alter_name4 + 1

            # 5 PERIODS

            for i in x:
                # Делим моменты
                index = x.index(i)
                for j in x[1:]:
                    index2 = x.index(j)
                    for k in x[2:]:
                        index3 = x.index(k)
                        for l in x[3:]:
                            index4 = x.index(l)
                            x_list_1 = x[:index + 1]
                            x_list_2 = x[index + 1:index2 + 1]
                            x_list_3 = x[index2 + 1:index3 + 1]
                            x_list_4 = x[index3 + 1:index4 + 1]
                            x_list_5 = x[index4 + 1:]

                            y_list_1 = y[:index + 1]
                            y_list_2 = y[index + 1:index2 + 1]
                            y_list_3 = y[index2 + 1:index3 + 1]
                            y_list_4 = y[index3 + 1:index4 + 1]
                            y_list_5 = y[index4 + 1:]

                            # Проверяем правильность разделения
                            intersection = set(y_list_1) & set(y_list_2)
                            intersection2 = set(y_list_2) & set(y_list_3)
                            intersection3 = set(y_list_3) & set(y_list_4)
                            intersection4 = set(y_list_4) & set(y_list_5)

                            # Запись в датафрейм
                            if not intersection and not intersection2 and not intersection3 and not intersection4 and y_list_1 and y_list_2 and y_list_3 and y_list_4 and y_list_5 and len(
                                    x_list_1) < 4 and len(x_list_2) < 4 and len(x_list_3) < 4 and len(
                                x_list_4) < 4 and len(x_list_5) < 4:
                                exemplares_5.append(
                                    {'diagnos': diagnos, 'name': f'альтернатива_5.{str(alter_name5)}', 'MH': f'{mh}',
                                     'tag': f'{tags}',
                                     'NPD': 5, 'Values_PD_1': set(y_list_1), 'lower_bound_1': int(x_list_1[0]),
                                     'upper_bound_1': int(x_list_1[-1]), 'Values_PD_2': set(y_list_2),
                                     'lower_bound_2': int(int(x_list_2[0]) - int(x_list_1[-1])),
                                     'upper_bound_2': int(int(x_list_2[-1]) - int(x_list_1[-1])),
                                     'Values_PD_3': set(y_list_3),
                                     'lower_bound_3': int(int(x_list_3[0]) - int(x_list_2[-1])),
                                     'upper_bound_3': int(int(x_list_3[-1]) - int(x_list_2[-1])),
                                     'Values_PD_4': set(y_list_4),
                                     'lower_bound_4': int(int(x_list_4[0]) - int(x_list_3[-1])),
                                     'upper_bound_4': int(int(x_list_4[-1]) - int(x_list_3[-1])),
                                     'Values_PD_5': set(y_list_5),
                                     'lower_bound_5': int(int(x_list_5[0]) - int(x_list_4[-1])),
                                     'upper_bound_5': int(int(x_list_5[-1]) - int(x_list_4[-1]))})
                                alter_name5 = alter_name5 + 1

def Merger_alternatives(exemplares, periods):
    test_exempl = exemplares
    Combinations = []
    all_combinations = []
    deleted = []
    new_alternatives = []
    final_alternatives = []

    # Будем перебирать ИБ 2 раза - сначала одного диагноза, затем второго
    mh = set(pd.DataFrame(exemplares).MH.drop_duplicates())
    mh_diagnos_1 = {'ИБ1', 'ИБ2', 'ИБ3', 'ИБ4'}
    mh_diagnos_2 = {'ИБ5', 'ИБ6', 'ИБ7', 'ИБ8'}
    mh_diagnos_1.intersection(mh)
    mh_diagnos_2.intersection(mh)
    mh_diagnos_1 = list(mh_diagnos_1)
    mh_diagnos_2 = list(mh_diagnos_2)
    mh = []
    mh.append(mh_diagnos_1)
    mh.append(mh_diagnos_2)

    for mh in mh:
        new_alternatives = []
        for I in range(len(mh) - 1):
            if len(new_alternatives) == 0:
                # задаем значение, по которому нужно отфильтровать словари
                filter_value_1 = f"{mh[I]}"
                filter_value_2 = f"{mh[I + 1]}"
                # фильтруем список словарей
                filtered_list_1 = [d for d in test_exempl if d.get('MH') == filter_value_1]
                filtered_list_2 = [d for d in test_exempl if d.get('MH') == filter_value_2]

            else:
                # Далее вторым словарем для объединения выбираем альтернативы
                filter_value_2 = f"{mh[I + 1]}"
                filtered_list_1 = new_alternatives
                filtered_list_2 = [d for d in test_exempl if d.get('MH') == filter_value_2]

            # Проверим на совместимость первй список со вторым
            new_alternatives = []
            for i in filtered_list_1:
                add = 0
                for j in filtered_list_2:
                    if i.get('tag') == j.get('tag') and i.get('diagnos') == j.get('diagnos'):
                        match periods:
                            case '1':
                                new_alternatives.append({
                                    'diagnos': i.get('diagnos'), 'name': f'комбинация',
                                    'MH': f'{i.get("MH")}_{j.get("MH")}',
                                    'tag': f"{i.get('tag')}", 'NPD': 1,
                                    'Values_PD': set(i.get('Values_PD')).union(set(j.get('Values_PD'))),
                                    'lower_bound': min(int(i.get('lower_bound')), int(j.get('lower_bound'))),
                                    'upper_bound': max(int(i.get('upper_bound')), int(j.get('upper_bound')))
                                })
                            case '2':
                                new_alternatives.append({
                                    'diagnos': i.get('diagnos'), 'name': f'комбинация',
                                    'MH': f'{i.get("MH")}_{j.get("MH")}',
                                    'tag': f"{i.get('tag')}", 'NPD': 2,
                                    'Values_PD_1': set(i.get('Values_PD_1')).union(set(j.get('Values_PD_1'))),
                                    'lower_bound_1': min(int(i.get('lower_bound_1')), int(j.get('lower_bound_1'))),
                                    'upper_bound_1': max(int(i.get('upper_bound_1')), int(j.get('upper_bound_1'))),
                                    'Values_PD_2': set(i.get('Values_PD_2')).union(set(j.get('Values_PD_2'))),
                                    'lower_bound_2': min(int(i.get('lower_bound_2')), int(j.get('lower_bound_2'))),
                                    'upper_bound_2': max(int(i.get('upper_bound_2')), int(j.get('upper_bound_2')))
                                })

                            case '3':
                                new_alternatives.append({
                                    'diagnos': i.get('diagnos'), 'name': f'комбинация',
                                    'MH': f'{i.get("MH")}_{j.get("MH")}',
                                    'tag': f"{i.get('tag')}", 'NPD': 3,
                                    'Values_PD_1': set(i.get('Values_PD_1')).union(set(j.get('Values_PD_1'))),
                                    'lower_bound_1': min(int(i.get('lower_bound_1')), int(j.get('lower_bound_1'))),
                                    'upper_bound_1': max(int(i.get('upper_bound_1')), int(j.get('upper_bound_1'))),
                                    'Values_PD_2': set(i.get('Values_PD_2')).union(set(j.get('Values_PD_2'))),
                                    'lower_bound_2': min(int(i.get('lower_bound_2')), int(j.get('lower_bound_2'))),
                                    'upper_bound_2': max(int(i.get('upper_bound_2')), int(j.get('upper_bound_2'))),
                                    'Values_PD_3': set(i.get('Values_PD_3')).union(set(j.get('Values_PD_3'))),
                                    'lower_bound_3': min(int(i.get(f'lower_bound_3')), int(j.get('lower_bound_3'))),
                                    'upper_bound_3': max(int(i.get('upper_bound_3')), int(j.get('upper_bound_3')))
                                })

                            case '4':
                                new_alternatives.append({
                                    'diagnos': i.get('diagnos'), 'name': f'комбинация',
                                    'MH': f'{i.get("MH")}_{j.get("MH")}',
                                    'tag': f"{i.get('tag')}", 'NPD': 4,
                                    'Values_PD_1': set(i.get('Values_PD_1')).union(set(j.get('Values_PD_1'))),
                                    'lower_bound_1': min(int(i.get('lower_bound_1')), int(j.get('lower_bound_1'))),
                                    'upper_bound_1': max(int(i.get('upper_bound_1')), int(j.get('upper_bound_1'))),
                                    'Values_PD_2': set(i.get('Values_PD_2')).union(set(j.get('Values_PD_2'))),
                                    'lower_bound_2': min(int(i.get('lower_bound_2')), int(j.get('lower_bound_2'))),
                                    'upper_bound_2': max(int(i.get('upper_bound_2')), int(j.get('upper_bound_2'))),
                                    'Values_PD_3': set(i.get('Values_PD_3')).union(set(j.get('Values_PD_3'))),
                                    'lower_bound_3': min(int(i.get('lower_bound_3')), int(j.get('lower_bound_3'))),
                                    'upper_bound_3': max(int(i.get('upper_bound_3')), int(j.get('upper_bound_3'))),
                                    'Values_PD_4': set(i.get('Values_PD_4')).union(set(j.get('Values_PD_4'))),
                                    'lower_bound_4': min(int(i.get('lower_bound_4')), int(j.get('lower_bound_4'))),
                                    'upper_bound_4': max(int(i.get('upper_bound_4')), int(j.get('upper_bound_4')))
                                })

                            case '5':
                                new_alternatives.append({
                                    'diagnos': i.get('diagnos'), 'name': f'комбинация',
                                    'MH': f'{i.get("MH")}_{j.get("MH")}',
                                    'tag': f"{i.get('tag')}", 'NPD': 5,
                                    'Values_PD_1': set(i.get('Values_PD_1')).union(set(j.get('Values_PD_1'))),
                                    'lower_bound_1': min(int(i.get('lower_bound_1')), int(j.get('lower_bound_1'))),
                                    'upper_bound_1': max(int(i.get('upper_bound_1')), int(j.get('upper_bound_1'))),
                                    'Values_PD_2': set(i.get('Values_PD_2')).union(set(j.get('Values_PD_2'))),
                                    'lower_bound_2': min(int(i.get('lower_bound_2')), int(j.get('lower_bound_2'))),
                                    'upper_bound_2': max(int(i.get('upper_bound_2')), int(j.get('upper_bound_2'))),
                                    'Values_PD_3': set(i.get('Values_PD_3')).union(set(j.get('Values_PD_3'))),
                                    'lower_bound_3': min(int(i.get('lower_bound_3')), int(j.get('lower_bound_3'))),
                                    'upper_bound_3': max(int(i.get('upper_bound_3')), int(j.get('upper_bound_3'))),
                                    'Values_PD_4': set(i.get('Values_PD_4')).union(set(j.get('Values_PD_4'))),
                                    'lower_bound_4': min(int(i.get('lower_bound_4')), int(j.get('lower_bound_4'))),
                                    'upper_bound_4': max(int(i.get('upper_bound_4')), int(j.get('upper_bound_4'))),
                                    'Values_PD_5': set(i.get('Values_PD_5')).union(set(j.get('Values_PD_5'))),
                                    'lower_bound_5': min(int(i.get('lower_bound_5')), int(j.get('lower_bound_5'))),
                                    'upper_bound_5': max(int(i.get('upper_bound_5')), int(j.get('upper_bound_5')))
                                })

                        add = add + 1

                # если альтернатива ни с чем не объединилась - добавить ее саму
                if add == 0:
                    new_alternatives.append(i)
                    # print(i.get('name'), ', мое бедное дитя, не нашла себе пару!!!!!!')
                # else:
                # print(i.get('name'), ' нашла себе пару')

            # Проверим второй список
            for i in filtered_list_2:
                add_1 = 0
                for j in filtered_list_1:
                    if i.get('tag') == j.get('tag') and i.get('diagnos') == j.get('diagnos'):
                        add_1 = add_1 + 1
                if add_1 == 0:
                    new_alternatives.append(i)
                    # print(i.get('name'), ', мое бедное дитя, не нашла себе пару!!!!!!')

            # Все альтернативы
            all_combinations += new_alternatives
            # Удалить дубли, если они вылезли
            all_combinations_ = all_combinations
            without_dublicates = []
            for elem in all_combinations_:
                if elem not in without_dublicates:
                    without_dublicates.append(elem)
            all_combinations = without_dublicates

            # Проверить, если ли пересечения периодов или значения больше 3
            if periods != '1':
                to_delete = set()
                for i, d in enumerate(new_alternatives):
                    match periods:
                        case '2':
                            values = set(d['Values_PD_1']).intersection(set(d['Values_PD_2']))

                        case '3':
                            values = set(d['Values_PD_1']).intersection(set(d['Values_PD_2'])) | set(
                                d['Values_PD_2']).intersection(set(
                                d['Values_PD_3']))
                        case '4':
                            values = set(d['Values_PD_1']).intersection(set(d['Values_PD_2'])) | set(
                                d['Values_PD_2']).intersection(set(
                                d['Values_PD_3'])) | set(d['Values_PD_3']).intersection(set(d['Values_PD_4']))

                        case '5':
                            values = set(d['Values_PD_1']).intersection(set(d['Values_PD_2'])) | set(
                                d['Values_PD_2']).intersection(set(
                                d['Values_PD_3'])) | set(d['Values_PD_3']).intersection(set(d['Values_PD_4'])) | set(
                                d['Values_PD_4']).intersection(set(d['Values_PD_5']))
                    if values:
                        to_delete.add(i)
                # Алтернативы, которые удалялись из-за пересечения значения периодов
                deleted += [d for i, d in enumerate(new_alternatives) if i in to_delete]
                new_alternatives = [d for i, d in enumerate(new_alternatives) if i not in to_delete]
                # Альтернативы, которые были за все шаги (после удаления)
                Combinations += new_alternatives
        final_alternatives += new_alternatives
    return all_combinations, final_alternatives

def Save_alternatives(name, all_alternatives, all_combinations, final_alternatives):
    wb = Workbook()
    ws = wb.active

    df1 = pd.DataFrame(all_alternatives)
    df2 = pd.DataFrame(all_combinations)
    df3 = pd.DataFrame(final_alternatives)

    df3.to_excel(writer, sheet_name=f'{name}', startrow=0, index=False)
    df1.to_excel(writer, sheet_name=f'{name}', startrow=len(df3) + 3, index=False, header=None)
    df2.to_excel(writer, sheet_name=f'{name}', startrow=len(df3) + len(df1) + 6, index=False, header=None)

    # Выделение зеленым цветом
    workbook = writer.book
    worksheet = writer.sheets[f'{name}']
    green_format = workbook.add_format({'bg_color': '#ccf598', 'font_color': '#050505'})
    for i in range(1, len(df3) + 1):
        worksheet.set_row(i, cell_format=green_format)

def Analis_alternatives(one, two, three, four, five, name):
    df1 = pd.DataFrame(one)
    df1 = df1.groupby('diagnos').get_group(f'{name}')
    df2 = pd.DataFrame(two)
    df2 = df2.groupby('diagnos').get_group(f'{name}')
    df3 = pd.DataFrame(three)
    df3 = df3.groupby('diagnos').get_group(f'{name}')
    df4 = pd.DataFrame(four)
    df4 = df4.groupby('diagnos').get_group(f'{name}')
    df5 = pd.DataFrame(five)
    df5 = df5.groupby('diagnos').get_group(f'{name}')

    df1.to_excel(writer, sheet_name=f'{name}', startrow=0, index=False)
    df2.to_excel(writer, sheet_name=f'{name}', startrow=len(df1) + 2, index=False, header=None)
    df3.to_excel(writer, sheet_name=f'{name}', startrow=len(df1) + len(df2) + 3, index=False, header=None)
    df4.to_excel(writer, sheet_name=f'{name}', startrow=len(df1) + len(df2) + len(df3) + 4, index=False, header=None)
    df5.to_excel(writer, sheet_name=f'{name}', startrow=len(df1) + len(df2) + len(df3) + len(df4) + 5, index=False,
                 header=None)

do_DB_Alternatives(df)

one_p = Merger_alternatives(exemplares_1, '1')
two_p = Merger_alternatives(exemplares_2, '2')
three_p = Merger_alternatives(exemplares_3, '3')
four_p = Merger_alternatives(exemplares_4, '4')
five_p = Merger_alternatives(exemplares_5, '5')

# writer = pd.ExcelWriter('Задание_3.xlsx', engine='xlsxwriter')
# Save_alternatives('ЧПД_1', exemplares_1, one_p[0], one_p[1])
# Save_alternatives('ЧПД_2', exemplares_2, two_p[0], two_p[1])
# Save_alternatives('ЧПД_3', exemplares_3, three_p[0], three_p[1])
# Save_alternatives('ЧПД_4', exemplares_4, four_p[0], four_p[1])
# Save_alternatives('ЧПД_5', exemplares_5, five_p[0], five_p[1])
# writer.save()
#
# writer = pd.ExcelWriter('Задание_4.xlsx', engine='xlsxwriter')
# Analis_alternatives(one_p[1], two_p[1], three_p[1], four_p[1], five_p[1], 'Ангина')
# Analis_alternatives(one_p[1], two_p[1], three_p[1], four_p[1], five_p[1], 'Рак легких')
# writer.save()
