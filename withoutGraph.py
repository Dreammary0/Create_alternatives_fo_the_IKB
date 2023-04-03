import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook

# Считать файл
df = pd.read_excel("данные.xlsx", sheet_name='Данные')
print(df)

# Подключение списка признаков и ИБ
tag = set(df.Признак.drop_duplicates())
MH = set(df.ИБ.drop_duplicates())
print(MH,tag)


# Перебираем истории болезни
alter_name1=1
alter_name2=1
alter_name3=1
alter_name4=1
alter_name5=1

exemplares_1=[]
exemplares_2=[]
exemplares_3=[]
exemplares_4=[]
exemplares_5=[]


for mh in MH:
    result = df.loc[df['ИБ'] == mh]
    for tags in tag:

        # 1 and 2 periods
        result1=result.loc[result['Признак'] == tags]
        x= list(result1.Момент)
        y = list(result1.Значение)


        for i in x:
            # Делим моменты
            index = x.index(i)
            x_list_1 = x[:index+1]
            x_list_2 = x[index+1:]
            y_list_1 = y[:index+1]
            y_list_2 = y[index+1:]

            # Проверяем правильность разделения
            intersection = set(y_list_1) & set(y_list_2)
            if not intersection:
                if y_list_2:
 # 2 PERIODS
                # Запись в датафрейм
                    exemplares_2.append({'name':f'альтернатива_2.{str(alter_name2)}', 'MH':f'{mh}', 'tag':f'{tags}', 'NPD':'2', 'Values_PD_1':f'{set(y_list_1)}', 'lower_bound_1':f'{int(x_list_1[0])}', 'upper_bound_1':f'{int(x_list_1[-1])}', 'Values_PD_2': f'{set(y_list_2)}', 'lower_bound_2':f'{int(int(x_list_2[0]) - int(x_list_1[-1]))}', 'upper_bound_2':f'{int(int(x_list_2[-1])-int(x_list_1[-1]))}'})
                    alter_name2=alter_name2+1
                else:
# 1 PERIOD
                    # Запись в датафрейм
                    exemplares_1.append( {'name':f'альтернатива_1.{str(alter_name1)}', 'MH':f'{mh}','tag':f'{tags}', 'NPD':'1', 'Values_PD':f'{set(y_list_1)}', 'lower_bound':f'{int(x_list_1[0])}', 'upper_bound':f'{int(x_list_1[-1])}'})
                    alter_name1=alter_name1+1


# 3 PERIODS

        for i in x:
            # Делим моменты
            index = x.index(i)
            for j in x[1:]:
                index2=x.index(j)
                x_list_1 = x[:index+1]
                x_list_2 = x[index+1:index2+1]
                x_list_3 = x[index2+1:]
                y_list_1 = y[:index+1]
                y_list_2 = y[index+1:index2+1]
                y_list_3 = y[index2+1:]

                # Проверяем правильность разделения
                intersection = set(y_list_1) & set(y_list_2)
                intersection2 = set(y_list_2) & set(y_list_3)

                # Запись в датафрейм
                if not intersection and not intersection2 and y_list_1 and y_list_2 and y_list_3:
                    exemplares_3.append({'name':f'альтернатива_3.{str(alter_name2)}', 'MH':f'{mh}', 'tag':f'{tags}', 'NPD':'3', 'Values_PD_1':f'{set(y_list_1)}', 'lower_bound_1':f'{int(x_list_1[0])}', 'upper_bound_1':f'{int(x_list_1[-1])}', 'Values_PD_2': f'{set(y_list_2)}', 'lower_bound_2':f'{int(int(x_list_2[0]) - int(x_list_1[-1]))}', 'upper_bound_2':f'{int(int(x_list_2[-1])-int(x_list_1[-1]))}', 'Values_PD_3': f'{set(y_list_3)}', 'lower_bound_3':f'{int(int(x_list_3[0]) - int(x_list_2[-1]))}', 'upper_bound_3':f'{int(int(x_list_3[-1])-int(x_list_2[-1]))}'} )
                    alter_name3=alter_name3+1

# 4 PERIODS

        for i in x:
            # Делим моменты
            index = x.index(i)
            for j in x[1:]:
                index2=x.index(j)
                for k in x[2:]:
                    index3=x.index(k)
                    x_list_1 = x[:index+1]
                    x_list_2 = x[index+1:index2+1]
                    x_list_3 = x[index2+1:index3+1]
                    x_list_4 = x[index3+1:]

                    y_list_1 = y[:index+1]
                    y_list_2 = y[index+1:index2+1]
                    y_list_3 = y[index2+1:index3+1]
                    y_list_4 = y[index3+1:]


                    # Проверяем правильность разделения
                    intersection = set(y_list_1) & set(y_list_2)
                    intersection2 = set(y_list_2) & set(y_list_3)
                    intersection3 = set(y_list_3) & set(y_list_4)

                    # Запись в датафрейм
                    if not intersection and not intersection2 and not intersection3 and y_list_1 and y_list_2 and y_list_3 and y_list_4:
                        exemplares_4.append({'name':f'альтернатива_4.{str(alter_name4)}', 'MH':f'{mh}', 'tag':f'{tags}', 'NPD':'4', 'Values_PD_1':f'{set(y_list_1)}', 'lower_bound_1':f'{int(x_list_1[0])}', 'upper_bound_1':f'{int(x_list_1[-1])}', 'Values_PD_2': f'{set(y_list_2)}', 'lower_bound_2':f'{int(int(x_list_2[0]) - int(x_list_1[-1]))}', 'upper_bound_2':f'{int(int(x_list_2[-1])-int(x_list_1[-1]))}', 'Values_PD_3': f'{set(y_list_3)}', 'lower_bound_3':f'{int(int(x_list_3[0]) - int(x_list_2[-1]))}', 'upper_bound_3':f'{int(int(x_list_3[-1])-int(x_list_2[-1]))}' , 'Values_PD_4': f'{set(y_list_4)}', 'lower_bound_4':f'{int(int(x_list_4[0]) - int(x_list_3[-1]))}', 'upper_bound_4':f'{int(int(x_list_4[-1])-int(x_list_3[-1]))}' } )
                        alter_name4=alter_name4+1


# 5 PERIODS

        for i in x:
            # Делим моменты
            index = x.index(i)
            for j in x[1:]:
                index2=x.index(j)
                for k in x[2:]:
                    index3=x.index(k)
                    for l in x[3:]:
                        index4=x.index(l)
                        x_list_1 = x[:index+1]
                        x_list_2 = x[index+1:index2+1]
                        x_list_3 = x[index2+1:index3+1]
                        x_list_4 = x[index3+1:index4+1]
                        x_list_5 = x[index4+1:]

                        y_list_1 = y[:index+1]
                        y_list_2 = y[index+1:index2+1]
                        y_list_3 = y[index2+1:index3+1]
                        y_list_4 = y[index3+1:index4+1]
                        y_list_5 = y[index4+1:]

                        # Проверяем правильность разделения
                        intersection = set(y_list_1) & set(y_list_2)
                        intersection2 = set(y_list_2) & set(y_list_3)
                        intersection3 = set(y_list_3) & set(y_list_4)
                        intersection4 = set(y_list_4) & set(y_list_5)

                        # Запись в датафрейм
                        if not intersection and not intersection2 and not intersection3 and not intersection4 and y_list_1 and y_list_2 and y_list_3 and y_list_4 and y_list_5:
                            exemplares_5.append({'name':f'альтернатива_5.{str(alter_name5)}', 'MH':f'{mh}', 'tag':f'{tags}', 'NPD':'5', 'Values_PD_1':f'{set(y_list_1)}', 'lower_bound_1':f'{int(x_list_1[0])}', 'upper_bound_1':f'{int(x_list_1[-1])}', 'Values_PD_2': f'{set(y_list_2)}', 'lower_bound_2':f'{int(int(x_list_2[0]) - int(x_list_1[-1]))}', 'upper_bound_2':f'{int(int(x_list_2[-1])-int(x_list_1[-1]))}', 'Values_PD_3': f'{set(y_list_3)}', 'lower_bound_3':f'{int(int(x_list_3[0]) - int(x_list_2[-1]))}', 'upper_bound_3':f'{int(int(x_list_3[-1])-int(x_list_2[-1]))}' , 'Values_PD_4': f'{set(y_list_4)}', 'lower_bound_4':f'{int(int(x_list_4[0]) - int(x_list_3[-1]))}', 'upper_bound_4':f'{int(int(x_list_4[-1])-int(x_list_3[-1]))}', 'Values_PD_5': f'{set(y_list_5)}', 'lower_bound_5':f'{int(int(x_list_5[0]) - int(x_list_4[-1]))}', 'upper_bound_5':f'{int(int(x_list_5[-1])-int(x_list_4[-1]))}' } )
                            alter_name5=alter_name5+1


df_1 = pd.DataFrame(exemplares_1)
df_2 = pd.DataFrame(exemplares_2)
df_3 = pd.DataFrame(exemplares_3)
df_4 = pd.DataFrame(exemplares_4)
df_5 = pd.DataFrame(exemplares_5)

writer = pd.ExcelWriter('Задание_3.xlsx')

df_1.to_excel(writer, sheet_name='1 период', index=False)
df_2.to_excel(writer, sheet_name='2 периода', index=False)
df_3.to_excel(writer, sheet_name='3 периода', index=False)
df_4.to_excel(writer, sheet_name='4 периода', index=False)
df_5.to_excel(writer, sheet_name='5 периодов', index=False)

writer.save()
