from datetime import datetime
import pandas as pd
import numpy as np
import pickle
import json
import scipy
from scipy.optimize import minimize, fmin_powell, fmin, differential_evolution
import warnings
from sklearn.preprocessing import StandardScaler
from sklearn import preprocessing
pd.options.display.max_columns = 1000
pd.options.display.max_rows = 100
import os
import warnings
import keras
from keras.models import model_from_json
import openpyxl
from sklearn.externals import joblib
from scipy.optimize import brute
from keras import backend as K

ls_opt_need = [
    u'толщина стенки',
    u'диаметр',
    u'норм_1', u'норм_2', u'норм_3', u'норм_4',
    u'отпуск_1', u'отпуск_2', u'отпуск_3', u'отпуск_4',
    u'Углеродный коэффицент'
]
tit=[u'диаметр',
u'толщина стенки',
u'C',
u'Mn',
u'Si',
u'Cr',
u'Ni',
u'Cu',
u'Al',
u'длина (отпуск)',
u'длина (закалка)',
u'норм_1',
u'норм_2',
u'норм_3',
u'норм_4',
u'отпуск_1',
u'отпуск_2',
u'отпуск_3',
u'отпуск_4',
u'Углеродный коэффицент',
u'темп-ра терм. (норм.)',
u'темп-ра спрейр (норм.)',
u'скорость движения (норм.)',
u'темп-ра терм. (отпуск)',
u'темп-ра спрейр (отпуск)',
u'скорость движения (отпуск)',
u'расход воды (норм.) (1)',
u'расход воды (норм.) (2)',
u'расход воды (норм.) (3)',
u'удельный расход воды',
u'параметр отпуска',
u'параметр закалки']

ls_fit_param = [u'темп-ра терм. (норм.)',
                u'темп-ра спрейр (норм.)',
                u'скорость движения (норм.)',
                u'темп-ра терм. (отпуск)',
                u'темп-ра спрейр (отпуск)',
                u'скорость движения (отпуск)',
                u'расход воды (норм.) (1)',
                u'расход воды (норм.) (2)',
                u'расход воды (норм.) (3)']

# Среднии значения некоторых хим элеметов не принимающих участие в предсказании
Mo = 0.007
V = 0.053
W = 0
S = 0.004
N = 0.0062
Ti = 0.0017
Nb = 0.00357
B = 0.00094
P = 0.0083

def fill_up(x):
    if pd.isnull(x[u'Предел текучести верхняя граница']):
        x[u'Предел текучести верхняя граница'] = x[u'Предел текучести нижняя граница'] + 30
    if pd.isnull(x[u'Предел прочности верхняя граница']):
        x[u'Предел прочности верхняя граница'] = x[u'Предел прочности нижняя граница'] + 30
    return x

# Функции для заполнения пропусков
# Если пропущен химиечский элемент - добавляеться среднее значение

def auto_fill_C(x):
    if (x == 0) | (x == None) | (pd.isnull(x)):
        return 0.15
    return x

def auto_fill_Mn(x):
    if (x == 0) | (x == None)| (pd.isnull(x)):
        return 0.55
    return x

def auto_fill_Si(x):
    if (x == 0) | (x == None)| (pd.isnull(x)):
        return 0.26
    return x

def auto_fill_Cr(x):
    if (x == 0) | (x == None)| (pd.isnull(x)):
        return 0.55
    return x

def auto_fill_Ni(x):
    if (x == 0) | (x == None)| (pd.isnull(x)):
        return 0.13
    return x

def auto_fill_Cu(x):
    if (x == 0) | (x == None)| (pd.isnull(x)):
        return 0.22
    return x

def auto_fill_Al(x):
    if (x == 0) | (x == None)| (pd.isnull(x)):
        return 0.03
    return x

def cl_inf(x):
    if x==float('inf'):
        x = 0
    return x

def auto_fill_empty(df):
    df[u'C'] = df[u'C'].apply(auto_fill_C)
    df[u'Mn'] = df[u'Mn'].apply(auto_fill_Mn)
    df[u'Si'] = df[u'Si'].apply(auto_fill_Si)
    df[u'Cr'] = df[u'Cr'].apply(auto_fill_Cr)
    df[u'Ni'] = df[u'Ni'].apply(auto_fill_Ni)
    df[u'Cu'] = df[u'Cu'].apply(auto_fill_Cu)
    df[u'Al'] = df[u'Al'].apply(auto_fill_Al)
    return df

# функции для расчета точе AC1 и AC3 (нужны для определения границ температур) (решили не использовать)

def calc_AC3_1(df):
    df[u'AC3_1'] = (911 - df.C*370-df.Mn*27.4+27.3*df.Si-6.35*df.Cr-\
    32.7*df.Ni+95.2*V+70.2*Ti+72*df.Al+64.5*Nb+\
    332*S+276*P-485*N+16.2*df.C*df.Mn+32.3*df.C*df.Si+\
    15.4*df.C*df.Cr+48*df.C*df.Ni+4.8*df.Mn*df.Ni+4.32*df.Si*df.Ni-\
    17.3*df.Si*Mo-18.6*df.Si*df.Ni+40.5*Mo*V+174*df.C*df.C+\
    2.46*df.Mn*df.Mn-6.86*df.Si*df.Si+0.322*df.Cr*df.Cr+9.9*Mo*Mo+\
    1.24*df.Ni*df.Ni-60.2*V*V-\
    900*B+5.57*W).round(0)
    return df

def calc_AC3_2(df):
    df[u'AC3_2'] = (912 - df.C*370-df.Mn*27.4+27.3*df.Si-6.35*df.Cr-\
    32.7*df.Ni+95.2*V+190*Ti+72*df.Al+64.5*Nb+\
    332*S+276*P+485*N+16.2*df.C*df.Mn+32.3*df.C*df.Si+\
    15.4*df.C*df.Cr+48*df.C*df.Ni+4.32*df.Si*df.Cr-17.3*df.Si*Mo-18.6*df.Si*df.Ni+\
    4.8*df.Mn*df.Ni+40.5*Mo*V+174*df.C*df.C+2.46*df.Mn*df.Mn-6.86*df.Si*df.Si+0.322*df.Cr*df.Cr+\
    9.9*Mo*Mo+1.24*df.Ni*df.Ni-60.2*V*V-
    900*B+5.57*W).round(0)
    return df

def calc_AC1_1(df):
    df[u'AC1_1'] = (723-7.08*df.Mn+37.7*df.Si+18.1*df.Cr+44.2*Mo-8.95*df.Ni+50.1*V+21.7*df.Al+3.18*W+\
    297*S-830*N-11.5*df.C*df.Si-14*df.Mn*df.Si-3.1*df.Cr*df.Si-57.9*df.C*Mo-15.5*df.Mn*Mo-\
    5.28*df.C*df.Ni-6*df.Mn*df.Ni+6.77*df.Si*df.Ni-0.8*df.Cr*df.Ni-27.4*df.C*V+30.8*Mo*V-\
    0.84*df.Cr*df.Cr-3.46*Mo*Mo-0.46*df.Ni*df.Ni-28*V*V).round(0)
    return df

def calc_AC1_2(df):
    df[u'AC1_2'] = (723-7.08*df.Mn+37.7*df.Si+18.1*df.Cr+44.2*Mo+8.95*df.Ni+50.1*V+21.7*df.Al+3.18*W+\
    297*S-830*N-11.5*df.C*df.Si-14*df.Mn*df.Si-3.1*df.Cr*df.Si-57.9*df.C*Mo-15.5*df.Mn*Mo-\
    5.28*df.C*df.Ni-6*df.Mn*df.Ni+6.77*df.Si*df.Ni-0.8*df.Cr*df.Ni-27.4*df.C*V+30.8*Mo*V-\
    0.84*df.Cr*df.Cr-3.46*Mo*Mo-0.46*df.Ni*df.Ni-28*V*V).round(0)
    return df

def calc_AC(df):
    df = calc_AC3_1(df)
    df = calc_AC3_2(df)
    df = calc_AC1_1(df)
    df = calc_AC1_2(df)
    df[u'AC3'] = (df[u'AC3_1'] + df[u'AC3_2'])/2.
    df[u'AC1'] = (df[u'AC1_1'] + df[u'AC1_2'])/2.
    del df[u'AC3_1']
    del df[u'AC3_2']
    del df[u'AC1_1']
    del df[u'AC1_2']
    return df


# Загрузка списка признаков, scaler и моделей

def load_model():

    titles_non_cat_data = json.load(open('app/model/titles_non_cat.json', "r"))
    titles_cat_data = json.load(open('app/model/titles_cat.json', "r"))

    scaler_fluidity = preprocessing.StandardScaler()
    scale_data_fluidity = json.load(open('app/model/fluidity/scaler', "r"))
    scaler_fluidity.mean_ = scale_data_fluidity[0]
    scaler_fluidity.scale_ = scale_data_fluidity[1]

    scaler_strength = preprocessing.StandardScaler()
    scale_data_strength = json.load(open('app/model/strength/scaler', "r"))
    scaler_strength.mean_ = scale_data_strength[0]
    scaler_strength.scale_ = scale_data_strength[1]

    json_file_fluidity = open('app/model/fluidity/NN.json', "r")
    loaded_model_fluidity = json_file_fluidity.read()
    json_file_fluidity.close()
    model_fluidity = model_from_json(loaded_model_fluidity)
    model_fluidity.load_weights('app/model/fluidity/NN.h5')
    model_fluidity.compile(loss=keras.losses.mean_squared_error, metrics=[keras.metrics.mean_squared_error],
                           optimizer=keras.optimizers.SGD(lr=0.0001, momentum=0.9, decay=1e-6))

    json_file_strength = open('app/model/strength/NN.json', "r")
    loaded_model_strength = json_file_strength.read()
    json_file_strength.close()
    model_strength = model_from_json(loaded_model_fluidity)
    model_strength.load_weights('app/model/strength/NN.h5')
    model_strength.compile(loss=keras.losses.mean_squared_error, metrics=[keras.metrics.mean_squared_error],
                           optimizer=keras.optimizers.SGD(lr=0.0001, momentum=0.9, decay=1e-6))

    grid_search_fluidity = pickle.load(open('app/model/fluidity/grid_search.sav', 'rb'))
    grid_search_strength = pickle.load(open('app/model/strength/grid_search.sav', 'rb'))

    GB_fluidity = pickle.load(open('app/model/fluidity/GB.sav', 'rb'))
    GB_strength = pickle.load(open('app/model/strength/GB.sav', 'rb'))

    return titles_non_cat_data, titles_cat_data, scaler_fluidity, scaler_strength, model_fluidity, model_strength, grid_search_fluidity, grid_search_strength, GB_fluidity, GB_strength


# Ищем похожий режим в исторических данных, от этого режима будем находить минимум функции
# Поиск идет по толщине стенки, затем по диаметру, затем по номеру ОКБ, затем по углеродному коэффиценту
# Выбрал углеродный коэффицент, так как признак наиболее сильно коррелируется с целевыми признаками

def close_value(database, col, value):
    database[u'diff'] = np.abs(database[col] - value)
    return database[database[u'diff'] == min(database[u'diff'])][col].values[0]


def find_row_close_sort(database, row, ls_need_col):
    for col in ls_opt_need:
        tmp = database[database[col] == row[col]]
        if tmp.shape[0] > 0:
            database = tmp
        else:
            value = close_value(database, col, row[col])
            database = database[database[col] == value]

    database = database.dropna(subset=ls_need_col)
    row_new = pd.Series(database.iloc[0, :])
    row_new[ls_opt_need] = row[ls_opt_need].copy()
    return row_new


def find_close_sort(database, df, ls_need_col):
    df = df.apply(lambda x: find_row_close_sort(database, x, ls_need_col), axis=1)
    return df

def get_index_diff(df1, df2):
    return list(set(df1.index).difference(set(df2.index)))


# Добавляем некоторые признаки

def add_param_OKB_and_Ccoef(data):
    data[u'длина (отпуск)'] = None
    data.loc[data[u'ОКБ (отпуск)'] == 1.0, u'длина (отпуск)'] = 2.2
    data.loc[data[u'ОКБ (отпуск)'] == 2.0, u'длина (отпуск)'] = 2.2
    data.loc[data[u'ОКБ (отпуск)'] == 3.0, u'длина (отпуск)'] = 1.2
    data.loc[data[u'ОКБ (отпуск)'] == 4.0, u'длина (отпуск)'] = 2.9
    data[u'длина (закалка)'] = None
    data.loc[data[u'ОКБ (закалка)'] == 1.0, u'длина (закалка)'] = 2.2
    data.loc[data[u'ОКБ (закалка)'] == 2.0, u'длина (закалка)'] = 2.2
    data.loc[data[u'ОКБ (закалка)'] == 3.0, u'длина (закалка)'] = 2.2
    data.loc[data[u'ОКБ (закалка)'] == 4.0, u'длина (закалка)'] = 1.65

    data[u'норм_1'] = 0
    data[u'норм_2'] = 0
    data[u'норм_3'] = 0
    data[u'норм_4'] = 0
    data[u'отпуск_1'] = 0
    data[u'отпуск_2'] = 0
    data[u'отпуск_3'] = 0
    data[u'отпуск_4'] = 0

    data.loc[data[u'ОКБ (отпуск)'] == 1.0, u'отпуск_1'] = 1
    data.loc[data[u'ОКБ (отпуск)'] == 2.0, u'отпуск_2'] = 1
    data.loc[data[u'ОКБ (отпуск)'] == 3.0, u'отпуск_3'] = 1
    data.loc[data[u'ОКБ (отпуск)'] == 4.0, u'отпуск_4'] = 1

    data.loc[data[u'ОКБ (закалка)'] == 1.0, u'норм_1'] = 1
    data.loc[data[u'ОКБ (закалка)'] == 2.0, u'норм_2'] = 1
    data.loc[data[u'ОКБ (закалка)'] == 3.0, u'норм_3'] = 1
    data.loc[data[u'ОКБ (закалка)'] == 4.0, u'норм_4'] = 1
    data[u'Углеродный коэффицент'] = data[u'C'] + data[u'Mn'] / 6 + (data[u'Cr']) / 5 + (data[u'Ni'] + data[u'Cu']) / 15
    return data

# Добавляем еще некоторые признаки

def add_param(data):
    data[u'длина (отпуск)'] = data[u'длина (отпуск)'].astype(float)
    data[u'длина (закалка)'] =data[u'длина (закалка)'].astype(float)
    data[u'скорость движения (норм.)'] = data[u'скорость движения (норм.)'].astype(float)
    data[u'скорость движения (отпуск)'] = data[u'скорость движения (отпуск)'].astype(float)
    data[u'расход воды (норм.)'] = data[u'расход воды (норм.) (1)']+data[u'расход воды (норм.) (2)']+ data[u'расход воды (норм.) (3)']
    data[u'удельный расход воды'] = data[u'расход воды (норм.)'] * 1000.0 / (data[u'диаметр'] * np.pi)
    data[u'параметр отпуска'] = (data[u'темп-ра терм. (отпуск)'] + 273.0) \
                         * (
                             20 + np.log(data[u'длина (отпуск)']) - np.log(
                                 data[u'скорость движения (отпуск)'] * 60.0)) \
                         * 1e-3
    data[u'параметр закалки'] = 1.0 / (1.0 / (data[u'темп-ра терм. (норм.)']+273.0) - 2.303 * 1.986 / 110000.0 * \
                                (np.log10(data[u'длина (закалка)']) - \
                                 np.log10(data[u'скорость движения (норм.)'] * 60.0))) - 273.0
    return data


# Превращение номера ОКБ из бинарного в категориальный признак для удобства представления в исходном файле

def OKB(data):
    data.loc[data[u'норм_1'] == 1.0, u'ОКБ (закалка)'] = 1
    data.loc[data[u'норм_2'] == 1.0, u'ОКБ (закалка)'] = 2
    data.loc[data[u'норм_3'] == 1.0, u'ОКБ (закалка)'] = 3
    data.loc[data[u'норм_4'] == 1.0, u'ОКБ (закалка)'] = 4

    data.loc[data[u'отпуск_1'] == 1.0, u'ОКБ (отпуск)'] = 1
    data.loc[data[u'отпуск_2'] == 1.0, u'ОКБ (отпуск)'] = 2
    data.loc[data[u'отпуск_3'] == 1.0, u'ОКБ (отпуск)'] = 3
    data.loc[data[u'отпуск_4'] == 1.0, u'ОКБ (отпуск)'] = 4

    return data

# Караем строки за нарушения и с клеймом отправялем в ошибочный файл

def err_output(df, raw_name):
    df.reset_index(inplace=True, drop=True)
    df[u'temp'] = df.index
    err_df = pd.DataFrame(columns=df.columns)
    df_corr = df.copy()
    for i in range(df.shape[0]):
        if df[df.index == i][raw_name].isnull().values.any():
            err_df = err_df.append(df[df.index==i],sort=False)
            err_df[u'Комментарий'] = u'отсутсвия значения: %s'%(raw_name)
            df_corr = df_corr.drop(df_corr[df_corr['temp']==i].index)
    del df_corr[u'temp']
    del err_df[u'temp']
    return [df_corr, err_df]

def verification(df):
    df[u'Номер строки'] = df.index+1
    df, err = err_output(df, u'Предел текучести нижняя граница')
    err_df = err
    df, err = err_output(df, u'Предел прочности нижняя граница')
    err_df = pd.concat([err_df, err])
    df, err = err_output(df, u'ОКБ (закалка)')
    err_df = pd.concat([err_df, err])
    df, err = err_output(df, u'ОКБ (отпуск)')
    err_df = pd.concat([err_df, err])
    return [df, err_df]

def get_index_diff(df1,df2):
    return list(set(df1.index).difference(set(df2.index)))

# Границы признаков, штрафуем функцию если выходит за границы

def param_min(data):
    data['темп-ра терм. (норм.)'] = 830
#     data[u'темп-ра терм. (норм.)'] = data[u'AC3']
    data[u'темп-ра спрейр (норм.)'] = 750
    data[u'скорость движения (норм.)'] = 0.3
    data[u'темп-ра терм. (отпуск)'] = 540
    data[u'темп-ра спрейр (отпуск)'] = 560
    data[u'скорость движения (отпуск)'] = 0.3
    data[u'расход воды (норм.) (1)'] = 10
    data[u'расход воды (норм.) (2)'] = 10
    data[u'расход воды (норм.) (3)'] = 0
    data[u'Предел текучести'] = data[u'Предел текучести нижняя граница']+3
    data[u'Предел прочности'] = data[u'Предел прочности нижняя граница']+3
    data = add_param(data)
    return data

def param_max(data):
    data[u'темп-ра терм. (норм.)'] = 980
    data[u'темп-ра спрейр (норм.)'] = 910
    data[u'скорость движения (норм.)'] = 3
    data['темп-ра терм. (отпуск)'] = 760
#     data[u'темп-ра терм. (отпуск)'] = data[u'AC1']
    data[u'темп-ра спрейр (отпуск)'] = 770
    data[u'скорость движения (отпуск)'] = 3
    data[u'расход воды (норм.) (1)'] = 100
    data[u'расход воды (норм.) (2)'] = 100
    data[u'расход воды (норм.) (3)'] = 70
    data[u'Предел прочности'] = data[u'Предел прочности верхняя граница']-3
    data[u'Предел текучести'] = data[u'Предел текучести верхняя граница']-3
    data = add_param(data)
    return data

# штраф за выход за границы

def is_in_bounds(row, bounds):
    penalty = 0
    for i in range(len(bounds)):
        if bounds[i][0]<=row.iloc[:, i].values<=bounds[i][1]:
            pass
        else:
            penalty += 200
    return penalty

# штраф за приближение к граница прочности и текучести
# границы взял +-3 потому что макимальная среднняя абсолютная ошибка примерно 2.5 и почему бы и нет

def bounds_target(pred_ys, pred_h ,bounds):
    penalty = 0
    if bounds[0][0]<=pred_ys<=bounds[0][1]:
        pass
    else:
        penalty += 300
    if bounds[1][0]<=pred_h<=bounds[1][1]:
        pass
    else:
        penalty += 300
    return penalty

# Предсказываем целевые признаки
# Предсказываем значение по трем моделям и находим средний ответ (просто так точнее немного выходит)

def pred_param(model, grid_search, GB, scaler, test, titles_non_cat_data, titles_cat_data):
    sc_data = test[titles_non_cat_data]
    sc_data = scaler.transform(sc_data)
    ct_data = test[titles_cat_data]
    sc_data = pd.DataFrame(sc_data, index=ct_data.index)
    inputs = sc_data.combine_first(ct_data).values

    pred1 = model.predict(inputs)
    pred2 = grid_search.predict(test[titles_non_cat_data + titles_cat_data])
    pred3 = GB.predict(test[titles_non_cat_data + titles_cat_data])
    pred = (pred1[:, 0] + pred2 + pred3) / 3
    return pred

# Функция, которую минимизируем
# штрафуем за выход признаков и целевых значений за границы, за низкую скорость, за отход целевых значений от среднего
def model_pr(fit_params, all_params, bounds, eta, models, grid_searchs, GBs, ls_need_cols, scalers, titles_non_cat_data, titles_cat_data):
    centr_ys = all_params[u'Текучесть середина']
    centr_h = all_params[u'Прочность середина']

    all_params = pd.concat([all_params[list(set(tit)-set(ls_fit_param))],
                            pd.Series(fit_params, index=ls_fit_param)])
    fit_params = pd.DataFrame(fit_params, index=ls_fit_param).T
    score = 0
# пересчет параметров
    all_params = pd.DataFrame(all_params).T
    all_params = add_param(all_params)

    all_params.reset_index(inplace=True, drop=True)
    all_params.dropna(inplace=True)
    # all_params = all_params.astype(np.float32)
    for model_name, model, grid_search, GB,ls_need_col, scaler in zip(['ys','h'], models, grid_searchs,GBs, ls_need_cols, scalers):
        if model_name=='h':
            centr=centr_h
            pred_h = pred_param(model, grid_search,GB, scaler,all_params, titles_non_cat_data, titles_cat_data)
            tmp_score = np.abs(pred_h - centr)
            tmp_score2 = tmp_score[0]
        else:
            centr=centr_ys
            pred_ys = pred_param(model, grid_search,GB, scaler,all_params, titles_non_cat_data, titles_cat_data)
            tmp_score = np.abs(pred_ys - centr)
            tmp_score2 = tmp_score[0]
        if tmp_score2 < 2:
            tmp_score2=0
        if tmp_score2 > 5:
            tmp_score2 += 100
        score += tmp_score2
    score += bounds_target(pred_ys, pred_h ,bounds[-2:])
    score += max(np.abs(2.2 - all_params[u'скорость движения (норм.)'].values),
                 np.abs(2.2 - all_params[u'скорость движения (отпуск)'].values))*eta
    score += is_in_bounds(fit_params, bounds[:-2])
    return score

def main(file,current_user):
    # Загружаем input файл и преобразуем его

    table_for_optimize = pd.read_excel(file)
    now = datetime.now()
    time = "%d_%ddate %d_%d_%dtime" % (now.day, now.month, now.hour, now.minute, now.second)
    input_filename = os.getcwd() + '/app/INPUT/' + "optimizer_input_" + time + "_" + current_user + ".xlsx"
    table_for_optimize.to_excel(input_filename)

    table_for_optimize.loc[
        table_for_optimize['Предел текучести нижняя граница'] > 110, 'Предел текучести нижняя граница'] = \
    table_for_optimize['Предел текучести нижняя граница'] / 9.8
    table_for_optimize.loc[
        table_for_optimize['Предел текучести верхняя граница'] > 110, 'Предел текучести верхняя граница'] = \
    table_for_optimize['Предел текучести верхняя граница'] / 9.8
    table_for_optimize.loc[
        table_for_optimize['Предел прочности нижняя граница'] > 110, 'Предел прочности нижняя граница'] = \
    table_for_optimize['Предел прочности нижняя граница'] / 9.8
    table_for_optimize.loc[
        table_for_optimize['Предел прочности верхняя граница'] > 110, 'Предел прочности верхняя граница'] = \
    table_for_optimize['Предел прочности верхняя граница'] / 9.8

    error1 = ''
    error2 = ''
    table_for_optimize, err_table_for_optimize = verification(table_for_optimize)
    if table_for_optimize.shape[0] == 0:
        error1 = 'Все строки удалены'
        error2 = "Строка " + str(err_table_for_optimize['Номер строки'][0]) + " удалена из-за " + err_table_for_optimize['Комментарий'][0] + "; \n"
        for i in range(1, err_table_for_optimize.shape[0] - 1):
            error2 = error2 + "cтрока " + str(err_table_for_optimize['Номер строки'][i]) + " удалена из-за " + err_table_for_optimize['Комментарий'][i] + "; \n"
        if err_table_for_optimize.shape[0] > 1:
            error2 = error2 + "cтрока " + str(
                err_table_for_optimize['Номер строки'][err_table_for_optimize.shape[0] - 1]) + " удалена из-за " + err_table_for_optimize['Комментарий'][err_table_for_optimize.shape[0] - 1] + "\n"
        return error1, error2

    table_for_optimize = auto_fill_empty(table_for_optimize)
    table_for_optimize = table_for_optimize.apply(fill_up, axis=1)
    table_for_optimize[u'Прочность середина'] = (table_for_optimize[
                                                     u'Предел прочности нижняя граница'] + table_for_optimize[
                                                     u'Предел прочности верхняя граница']) / 2.0
    table_for_optimize[u'Текучесть середина'] = (table_for_optimize[
                                                     u'Предел текучести нижняя граница'] + table_for_optimize[
                                                     u'Предел текучести верхняя граница']) / 2.0
    table_for_optimize = add_param_OKB_and_Ccoef(table_for_optimize)

    # Загружаем список признаков, scaler и модели
    titles_non_cat_data, titles_cat_data, scaler_fluidity, scaler_strength, model_fluidity, model_strength, grid_search_fluidity, grid_search_strength, GB_fluidity, GB_strength = load_model()

    # Загружаем исторические данные

    database = pd.read_excel('app/data/historical_data.xlsx')
    database = database[titles_non_cat_data + titles_cat_data]

    # Ищем исторический близкий режим и преобразовываем

    answ = find_close_sort(database, table_for_optimize, titles_non_cat_data + titles_cat_data)

    answ = answ[[u'темп-ра терм. (норм.)',
                 u'темп-ра спрейр (норм.)', u'скорость движения (норм.)',
                 u'темп-ра терм. (отпуск)', u'темп-ра спрейр (отпуск)',
                 u'скорость движения (отпуск)', u'расход воды (норм.) (1)',
                 u'расход воды (норм.) (2)', u'расход воды (норм.) (3)']]
    answ = pd.concat([table_for_optimize, answ], axis=1)
    answ = add_param(answ)
    answ = calc_AC(answ)
    answ.reset_index(inplace=True, drop=True)

    # Границы

    X_c_down = answ.copy()
    X_c_up = answ.copy()
    X_c_down = param_min(X_c_down)
    X_c_up = param_max(X_c_up)

    # Чтобы потом красиво засунуть в for

    models = [model_fluidity, model_strength]
    grid_searchs = [grid_search_fluidity, grid_search_strength]
    GBs = [GB_fluidity, GB_strength]
    ls_need_cols = [titles_non_cat_data + titles_cat_data, titles_non_cat_data + titles_cat_data]
    scalers = [scaler_fluidity, scaler_strength]

    answ_n = answ.copy()

    answers_array = []
    X_a = X_c_down.copy()
    X_b = X_c_up.copy()
    answ_n.reset_index(inplace=True, drop=True)
    X_a.reset_index(inplace=True, drop=True)
    X_b.reset_index(inplace=True, drop=True)

    for it in X_a.index:
        bounds = [(i, j) for i, j in zip(X_a.loc[it, ls_fit_param + [u'Предел текучести', u'Предел прочности']],
                                         X_b.loc[it, ls_fit_param + ['Предел текучести', 'Предел прочности']])]
        all_params = answ_n.iloc[it, :]
        fit_p = answ_n.loc[it, ls_fit_param]

        # находим минимум
        a = fmin(lambda fit_params: model_pr(fit_params,
                                             all_params,
                                             bounds,
                                             40, models, grid_searchs, GBs, ls_need_cols, scalers, titles_non_cat_data, titles_cat_data),
                 fit_p, maxfun=3000)
        test = pd.concat([all_params[list(set(tit) - set(ls_fit_param))],
                          pd.Series(a, index=ls_fit_param)])
        test = pd.DataFrame(test).T
        test = add_param(test)
        h = pred_param(model_strength, grid_search_strength, GB_strength, scaler_strength, test, titles_non_cat_data, titles_cat_data)
        ys = pred_param(model_fluidity, grid_search_fluidity, GB_fluidity, scaler_fluidity, test, titles_non_cat_data, titles_cat_data)

        all_params = pd.concat([all_params[list(set(tit) - set(ls_fit_param))],
                                pd.Series(a, index=ls_fit_param)])
        answers_array.append(all_params)
        answers_array[-1] = pd.concat(
            [answers_array[-1], pd.Series([ys[0], h[0]], index=[u'pred Текучесть', u'pred Прочность'])])
        print(bounds)
        print(a)
        print('Номер: ', it)
        print('Текучесть: ')
        print('Середина: ',  answ_n[answ_n.index == it]['Текучесть середина'].value_counts().index[0] )
        print('Низ: ',  answ_n[answ_n.index == it]['Предел текучести нижняя граница'].value_counts().index[0] )
        print('Верх: ', answ_n[answ_n.index == it]['Предел текучести верхняя граница'].value_counts().index[0] )
        print(ys)

        print()
        print('Прочность: ')
        print('Середина: ',  answ_n[answ_n.index == it]['Прочность середина'].value_counts().index[0] )
        print('Низ: ',  answ_n[answ_n.index == it]['Предел прочности нижняя граница'].value_counts().index[0] )
        print('Верх: ', answ_n[answ_n.index == it]['Предел прочности верхняя граница'].value_counts().index[0] )
        print(h)

    opt_answer_end = pd.DataFrame(answers_array)
    opt_answer_end = OKB(opt_answer_end)
    opt_answer_end = opt_answer_end[
        ls_need_cols[0] + [u'ОКБ (закалка)', u'ОКБ (отпуск)'] + [u'pred Текучесть', u'pred Прочность']].copy()

    opt_answer_end = pd.concat([opt_answer_end, answ[[u'Номер строки', u'Предел текучести нижняя граница',
                                                      u'Предел текучести верхняя граница',
                                                      u'Предел прочности нижняя граница',
                                                      u'Предел прочности верхняя граница']]], axis=1)
    opt_answer_end = opt_answer_end[[u'Номер строки', u'диаметр', u'толщина стенки', u'темп-ра терм. (норм.)',
                                     u'темп-ра спрейр (норм.)', u'скорость движения (норм.)',
                                     u'темп-ра терм. (отпуск)', u'темп-ра спрейр (отпуск)',
                                     u'скорость движения (отпуск)', u'C', u'Mn', u'Si', u'Cr', u'Ni', u'Cu', u'Al',
                                     u'расход воды (норм.) (1)',
                                     u'расход воды (норм.) (2)', u'расход воды (норм.) (3)',
                                     u'ОКБ (закалка)',
                                     u'ОКБ (отпуск)', u'Предел текучести нижняя граница',
                                     u'Предел текучести верхняя граница', u'pred Текучесть',
                                     u'Предел прочности нижняя граница',
                                     u'Предел прочности верхняя граница', u'pred Прочность']].copy()
    K.clear_session()
    # сохраняем результат
    if err_table_for_optimize.shape[0] != 0:
        err_table_for_optimize = err_table_for_optimize.reset_index()
        print(err_table_for_optimize)
        print(err_table_for_optimize.shape[0])
        error2 = "Строка " + str(err_table_for_optimize['Номер строки'][0]) + " удалена из-за " + err_table_for_optimize['Комментарий'][0] + "; \n"
        for i in range(1, err_table_for_optimize.shape[0]-1):
            error2 = error2 + "cтрока " + str(err_table_for_optimize['Номер строки'][i]) + " удалена из-за " + err_table_for_optimize['Комментарий'][i] + "; \n"
        if err_table_for_optimize.shape[0] > 1:
            error2 = error2 + "cтрока " + str(err_table_for_optimize['Номер строки'][err_table_for_optimize.shape[0]-1]) + " удалена из-за " + err_table_for_optimize['Комментарий'][ err_table_for_optimize.shape[0]-1] + "\n"

    if opt_answer_end.shape[0] != 0:
        direct = 'app/OUTPUT/output_optimizer' + time +  "_" + current_user +'.xlsx'
        writer = pd.ExcelWriter(direct, engine='xlsxwriter')
        opt_answer_end.to_excel(writer, index=None)
        writer.close()

    return error1,error2




