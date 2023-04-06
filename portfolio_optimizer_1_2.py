# -*- coding: utf-8 -*-
"""
generic portfolio mean-variance optimization based on CLA
version:
    1.0, initial implementation
    1.1, add sample input to UI
    1.2, subtract cash return in sharpe ratio
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import warnings
import seaborn as sns
import getpass
import datetime
import PySimpleGUI as sg
import sys
import win32com.client
import os

sns.set_style('darkgrid')
warnings.filterwarnings("ignore")
plt.rcParams["figure.figsize"] = [15, 8]
global dir_project
userlogin = os.getlogin().lower()
dir_project = f'C:/Users/{userlogin}/jana.com.au/Quantitative Applications - Documents/General/Models & Vendors & Guides/Portfolio-Optimizer/'
class CLA:

    def __init__(self, mean, covar, lB, uB, rf):
        # Initialize the class
        self.mean = mean
        self.covar = covar
        self.lB = lB
        self.uB = uB
        self.rf = rf
        self.w = []  # solution
        self.l = []  # lambdas
        self.g = []  # gammas
        self.f = []  # free weights

    # ---------------------------------------------------------------
    def solve(self):
        # Compute the turning points,free sets and weights
        f, w = self.initAlgo()
        self.w.append(np.copy(w))  # store solution
        self.l.append(-np.inf)
        self.g.append(-np.inf)
        self.f.append(f[:])
        while True:
            # 1) case a): Bound one free weight
            l_in = -np.inf
            if len(f) > 1:
                covarF, covarFB, meanF, wB = self.getMatrices(f)
                covarF_inv = np.linalg.inv(covarF)
                j = 0
                for i in f:
                    l, bi = self.computeLambda(covarF_inv, covarFB, meanF, wB, j, [self.lB[i], self.uB[i]])
                    if l > l_in: l_in, i_in, bi_in = l, i, bi
                    j += 1
            # 2) case b): Free one bounded weight
            l_out = -np.inf
            if len(f) < self.mean.shape[0]:
                b = self.getB(f)
                for i in b:
                    covarF, covarFB, meanF, wB = self.getMatrices(f + [i])
                    covarF_inv = np.linalg.inv(covarF)
                    l, bi = self.computeLambda(covarF_inv, covarFB, meanF, wB, meanF.shape[0] - 1, \
                                               self.w[-1][i])
                    if (self.l[-1] == -np.inf or l < self.l[-1]) and l > l_out: l_out, i_out = l, i
            if (l_in == -np.inf or l_in < 0) and (l_out == -np.inf or l_out < 0):
                # 3) compute minimum variance solution
                self.l.append(0)
                covarF, covarFB, meanF, wB = self.getMatrices(f)
                covarF_inv = np.linalg.inv(covarF)
                meanF = np.zeros(meanF.shape)
            else:
                # 4) decide lambda
                if l_in > l_out:
                    self.l.append(l_in)
                    f.remove(i_in)
                    w[i_in] = bi_in  # set value at the correct boundary
                else:
                    self.l.append(l_out)
                    f.append(i_out)
                covarF, covarFB, meanF, wB = self.getMatrices(f)
                covarF_inv = np.linalg.inv(covarF)
            # 5) compute solution vector
            wF, g = self.computeW(covarF_inv, covarFB, meanF, wB)
            for i in range(len(f)): w[f[i]] = wF[i]
            self.w.append(np.copy(w))  # store solution
            self.g.append(g)
            self.f.append(f[:])
            if self.l[-1] == 0: break
        # 6) Purge turning points
        self.purgeNumErr(10e-10)
        self.purgeExcess()

    # ---------------------------------------------------------------
    def initAlgo(self):
        # Initialize the algo
        # 1) Form structured array
        a = np.zeros((self.mean.shape[0]), dtype=[('id', int), ('mu', float)])
        b = [self.mean[i][0] for i in range(self.mean.shape[0])]
        a[:] = list(zip(range(self.mean.shape[0]), b))  # fill structured array
        # 2) Sort structured array
        b = np.sort(a, order='mu')
        # 3) First free weight
        i, w = b.shape[0], np.copy(self.lB)
        while sum(w) < 1:
            i -= 1
            w[b[i][0]] = self.uB[b[i][0]]
        w[b[i][0]] += 1 - sum(w)
        return [b[i][0]], w

    # ---------------------------------------------------------------
    def computeBi(self, c, bi):
        if c > 0:
            bi = bi[1][0]
        if c < 0:
            bi = bi[0][0]
        return bi

    # ---------------------------------------------------------------
    def computeW(self, covarF_inv, covarFB, meanF, wB):
        # 1) compute gamma
        onesF = np.ones(meanF.shape)
        g1 = np.dot(np.dot(onesF.T, covarF_inv), meanF)
        g2 = np.dot(np.dot(onesF.T, covarF_inv), onesF)
        if (wB is None) or (wB.sum() == 0):
            g, w1 = float(-self.l[-1] * g1 / g2 + 1 / g2), 0
        else:
            onesB = np.ones(wB.shape)
            g3 = np.dot(onesB.T, wB)
            g4 = np.dot(covarF_inv, covarFB)
            w1 = np.dot(g4, wB)
            g4 = np.dot(onesF.T, w1)
            g = float(-self.l[-1] * g1 / g2 + (1 - g3 + g4) / g2)
        # 2) compute weights
        w2 = np.dot(covarF_inv, onesF)
        w3 = np.dot(covarF_inv, meanF)
        return -w1 + g * w2 + self.l[-1] * w3, g

    # ---------------------------------------------------------------
    def computeLambda(self, covarF_inv, covarFB, meanF, wB, i, bi):
        # 1) C
        onesF = np.ones(meanF.shape)
        c1 = np.dot(np.dot(onesF.T, covarF_inv), onesF)
        c2 = np.dot(covarF_inv, meanF)
        c3 = np.dot(np.dot(onesF.T, covarF_inv), meanF)
        c4 = np.dot(covarF_inv, onesF)
        c = -c1 * c2[i] + c3 * c4[i]
        if c == 0: return -np.inf, -np.inf
        # 2) bi
        if type(bi) == list: bi = self.computeBi(c, bi)
        # 3) Lambda
        if (wB is None) or (wB.sum() == 0):
            # All free assets
            return float((c4[i] - c1 * bi) / c), bi

        else:
            onesB = np.ones(wB.shape)
            l1 = np.dot(onesB.T, wB)
            l2 = np.dot(covarF_inv, covarFB)
            l3 = np.dot(l2, wB)
            l2 = np.dot(onesF.T, l3)
            return float(((1 - l1 + l2) * c4[i] - c1 * (bi + l3[i])) / c), bi

    # ---------------------------------------------------------------
    def getMatrices(self, f):
        # Slice covarF,covarFB,covarB,meanF,meanB,wF,wB
        covarF = self.reduceMatrix(self.covar, f, f)
        meanF = self.reduceMatrix(self.mean, f, [0])
        b = self.getB(f)
        covarFB = self.reduceMatrix(self.covar, f, b)
        wB = self.reduceMatrix(self.w[-1].reshape(-1,1), b, [0])
        return covarF, covarFB, meanF, wB

    # ---------------------------------------------------------------
    def getB(self, f):
        return self.diffLists(range(self.mean.shape[0]), f)

    # ---------------------------------------------------------------
    def diffLists(self, list1, list2):
        return list(set(list1) - set(list2))

    # ---------------------------------------------------------------
    def reduceMatrix(self, matrix, listX, listY):
        # Reduce a matrix to the provided list of rows and columns
        if len(listX) == 0 or len(listY) == 0: return
        matrix_ = matrix[:, listY[0]:listY[0] + 1]
        for i in listY[1:]:
            a = matrix[:, i:i + 1]
            matrix_ = np.append(matrix_, a, 1)
        matrix__ = matrix_[listX[0]:listX[0] + 1, :]
        for i in listX[1:]:
            a = matrix_[i:i + 1, :]
            matrix__ = np.append(matrix__, a, 0)
        return matrix__

    # ---------------------------------------------------------------
    def purgeNumErr(self, tol):
        # Purge violations of inequality constraints (associated with ill-conditioned covar matrix)
        i = 0
        while True:
            flag = False
            if i == len(self.w): break
            if abs(sum(self.w[i]) - 1) > tol:
                flag = True
            else:
                for j in range(self.w[i].shape[0]):
                    if self.w[i][j] - self.lB[j] < -tol or self.w[i][j] - self.uB[j] > tol:
                        flag = True;
                        break
            if flag == True:
                del self.w[i]
                del self.l[i]
                del self.g[i]
                del self.f[i]
            else:
                i += 1
        return

    # ---------------------------------------------------------------
    def purgeExcess(self):
        # Remove violations of the convex hull
        i, repeat = 0, False
        while True:
            if repeat == False: i += 1
            if i == len(self.w) - 1: break
            w = self.w[i]
            mu = np.dot(w.T, self.mean)[0, 0]
            j, repeat = i + 1, False
            while True:
                if j == len(self.w): break
                w = self.w[j]
                mu_ = np.dot(w.T, self.mean)[0, 0]
                if mu < mu_:
                    del self.w[i]
                    del self.l[i]
                    del self.g[i]
                    del self.f[i]
                    repeat = True
                    break
                else:
                    j += 1
        return

    # ---------------------------------------------------------------
    def getMinVar(self):
        # Get the minimum variance solution
        var = []
        for w in self.w:
            a = np.dot(np.dot(w.T, self.covar), w)
            var.append(a)
        return min(var) ** .5, self.w[var.index(min(var))]

    # ---------------------------------------------------------------
    def getMaxSR(self):
        # Get the max Sharpe ratio portfolio
        # 1) Compute the local max SR portfolio between any two neighbor turning points
        w_sr, sr = [], []
        for i in range(len(self.w) - 1):
            w0 = np.copy(self.w[i])
            w1 = np.copy(self.w[i + 1])
            kargs = {'minimum': False, 'args': (w0, w1)}
            a, b = self.goldenSection(self.evalSR, 0, 1, **kargs)
            w_sr.append(a * w0 + (1 - a) * w1)
            sr.append(b)
        return max(sr), w_sr[sr.index(max(sr))]

    # ---------------------------------------------------------------
    def evalSR(self, a, w0, w1):
        # Evaluate SR of the portfolio within the convex combination
        w = a * w0 + (1 - a) * w1
        b = np.dot(w.T, self.mean)[0, 0]
        c = np.dot(np.dot(w.T, self.covar), w)[0, 0] ** .5
        return (b - self.rf) / c

    # ---------------------------------------------------------------
    def goldenSection(self, obj, a, b, **kargs):
        # Golden section method. Maximum if kargs['minimum']==False is passed
        from math import log, ceil
        tol, sign, args = 1.0e-9, 1, None
        if 'minimum' in kargs and kargs['minimum'] == False: sign = -1
        if 'args' in kargs: args = kargs['args']
        numIter = int(ceil(-2.078087 * log(tol / abs(b - a))))
        r = 0.618033989
        c = 1.0 - r
        # Initialize
        x1 = r * a + c * b;
        x2 = c * a + r * b
        f1 = sign * obj(x1, *args);
        f2 = sign * obj(x2, *args)
        # Loop
        for i in range(numIter):
            if f1 > f2:
                a = x1
                x1 = x2;
                f1 = f2
                x2 = c * a + r * b;
                f2 = sign * obj(x2, *args)
            else:
                b = x2
                x2 = x1;
                f2 = f1
                x1 = r * a + c * b;
                f1 = sign * obj(x1, *args)
        if f1 < f2:
            return x1, sign * f1
        else:
            return x2, sign * f2

    # ---------------------------------------------------------------
    def efFrontier(self, points):
        # Get the efficient frontier
        mu, sigma, weights = [], [], []
        a = np.linspace(0, 1, int(points / len(self.w)))[:-1]  # remove the 1, to avoid duplications
        b = range(len(self.w) - 1)
        for i in b:
            w0, w1 = self.w[i], self.w[i + 1]
            if i == b[-1]: a = np.linspace(0, 1, int(points / len(self.w)))  # include the 1 in the last iteration
            for j in a:
                w = w1 * j + (1 - j) * w0
                weights.append(np.copy(w))
                mu.append(np.dot(w.T, self.mean)[0, 0])
                sigma.append(np.dot(np.dot(w.T, self.covar), w)[0, 0] ** .5)
        return mu, sigma, weights


def plot2D(x, y, xLabel='', yLabel='', title='', pathChart=None, extra_point=None):
    import matplotlib.pyplot as mpl
    fig = mpl.figure()
    ax = fig.add_subplot(1, 1, 1)  # one row, one column, first plot
    ax.plot(x, y, color='blue')
    ax.set_xlabel(xLabel)
    ax.set_ylabel(yLabel, rotation=90)
    mpl.xticks(rotation='vertical')
    mpl.title(title)

    if extra_point != None:
        plt.plot(extra_point[0], extra_point[1],
                 'ro', label='r = {}, stdev = {}'.format(extra_point[1], extra_point[0]))
        ax.legend(loc='best')
    if pathChart == None:
        mpl.show()
    else:
        mpl.savefig(pathChart)
    mpl.clf()  # reset pylab

    return


def plot2D(x, y, xLabel='', yLabel='', title='', pathChart=None, extra_point=None):
    import matplotlib.pyplot as mpl
    fig = mpl.figure()
    ax = fig.add_subplot(1, 1, 1)  # one row, one column, first plot
    ax.plot(x, y, color='blue')
    ax.set_xlabel(xLabel)
    ax.set_ylabel(yLabel, rotation=90)
    mpl.xticks(rotation='vertical')
    mpl.title(title)

    if extra_point != None:
        plt.plot(extra_point[0], extra_point[1],
                 'ro', label='r = {}, stdev = {}'.format(extra_point[1], extra_point[0]))
        ax.legend(loc='best')
    if pathChart == None:
        mpl.show()
    else:
        mpl.savefig(pathChart)
    mpl.clf()  # reset pylab

    return

def clean_up_list(_list):

    return [str(x).strip() for x in _list]

def main(path_input, dir_out):

    ctime = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    # 1) Load data
    df_return_risk = pd.read_excel(path_input, sheet_name='ReturnRisk', index_col=0) \
                    .dropna(how='all', axis=0).dropna(how='all', axis=1)
    df_return_risk.index = clean_up_list(df_return_risk.index)
    list_portfolio_names = clean_up_list(df_return_risk['Portfolio_Name'].dropna().unique())
    list_asset = clean_up_list(df_return_risk.index.unique().to_numpy())
    rf = df_return_risk['Return'].loc['CSH'] if 'CSH' in list_asset else 0
    corr = pd.read_excel(path_input, sheet_name='Correlation', index_col=0).dropna(how='all')
    corr.index = clean_up_list(corr.index)
    corr.columns = clean_up_list(corr.columns)
    corr = corr[list_asset].loc[list_asset].to_numpy()
    for portfolio_name in list_portfolio_names:
        # 2) Loop through portfolios
        print(f"working on {portfolio_name}")
        df = df_return_risk[df_return_risk['Portfolio_Name'] == portfolio_name]
        mean = df['Return'].to_numpy().reshape(-1, 1)
        stdev = df['Volatility'].to_numpy()
        lower_bound = df['Min_Weight'].to_numpy().reshape(-1, 1) / 100
        upper_bound = df['Max_Weight'].to_numpy().reshape(-1, 1) / 100
        covar = np.diag(stdev) @ corr @ np.diag(stdev)
        # 3) Invoke object
        print("running optimization")
        cla = CLA(mean, covar, lower_bound, upper_bound, rf)
        cla.solve()
        # 4) Efficient Frontier
        print("creating efficient frontier")
        mu_ef, sigma_ef, w_ef = cla.efFrontier(100)
        df_weights = pd.DataFrame(np.array(w_ef).reshape(len(w_ef), len(mean)), columns=list_asset)
        # 5) Get Maximum Sharpe ratio portfolio
        sr, w_sr = cla.getMaxSR()
        mu_sr, sigma_sr = (w_sr.T @ mean)[0][0], (w_sr.T @ cla.covar @ w_sr)[0][0] ** 0.5
        # 6) Get Minimum Variance portfolio
        mv, w_mv = cla.getMinVar()
        mu_mv, sigma_mv = (w_mv.T @ mean)[0][0], (w_mv.T @ cla.covar @ w_mv)[0][0] ** 0.5
        # 7) Output
        df_weights = df_weights.append(pd.DataFrame(data=w_sr.T, index=['max_sharp'], columns=list_asset))
        df_weights = df_weights.append(pd.DataFrame(data=w_mv.T, index=['min_vol'], columns=list_asset))
        df_weights['volatility'] = df_weights.apply(lambda x: (np.array(x).T @ cla.covar @ np.array(x)) ** 0.5, axis=1)
        df_weights['return'] = (df_weights[list_asset] * mean.T).sum(axis=1)
        df_weights['sharp'] = (df_weights['return'] - rf) / df_weights['volatility']
        sns.lineplot(data=df_weights, x="volatility", y="return")
        plt.scatter(x=sigma_sr, y=mu_sr, color='g', marker='>', s=100)
        plt.text(sigma_sr, mu_sr, f'maximal Sharpe, r={round(mu_sr, 2)},vol={round(sigma_sr, 2)}')
        plt.scatter(x=sigma_mv, y=mu_mv, color='g', marker='>', s=100)
        plt.text(sigma_mv, mu_mv, f'minimal volatility, r={round(mu_mv, 2)},vol={round(sigma_mv, 2)}')
        # 8) Current Portfolio
        list_allocation = [x for x in df.columns if str(x).startswith('Allocation_')]
        for allocation in list_allocation:
            cur_w = df[allocation].to_numpy().reshape(-1, 1) / 100
            mu_cur, sigma_cur = (cur_w.T @ mean)[0][0], (cur_w.T @ cla.covar @ cur_w)[0][0] ** 0.5
            plt.scatter(x=sigma_cur, y=mu_cur, color='r', marker=',', s=100)
            plt.text(sigma_cur, mu_cur, f'{allocation}, r={round(mu_cur, 2)},vol={round(sigma_cur, 2)}')

        plt.title(f'{portfolio_name} Efficient Frontier')
        plt.xlabel('Volatility %')
        plt.ylabel('Return %')
        plt.tight_layout()
        plt.savefig(f'{dir_out}{getpass.getuser()}_{portfolio_name}_{ctime}.png')
        plt.close()
        df_weights.to_excel(f'{dir_out}{getpass.getuser()}_{portfolio_name}_{ctime}.xlsx')
    print('done')

    return

def run_portfolio_optimizer():

    """
    UI
    """
    sg.theme('DarkAmber')
    layout = [[sg.Text('Input File', size=(17, 1)), sg.InputText(), sg.FileBrowse(), sg.Button('Open Template')],
              [sg.Text('Output Directory', size=(17, 1)), sg.InputText(), sg.FolderBrowse()],
              [sg.Button('Ok'), sg.Button('Cancel')]]
    window = sg.Window('Portfolio Optimizer', layout, default_element_size=(50, 40), grab_anywhere=False)
    event, values = window.read()
    window.close()
    if event in (None, 'Cancel'):
        sys.exit("exit")
    if event == 'Open Template':
        print('opening example')
        excelApp = win32com.client.Dispatch('Excel.Application')
        excelApp.Workbooks.Open(dir_project + 'input/Input_Template.xlsx', None, True)
        excelApp.Visible = True
    else:
        main(str(values[0]), str(values[1])+'/')
    return

if __name__ == "__main__":

    path_input = 'input/Input_Template.xlsx'
    dir_out = 'output/'
    main(path_input, dir_out)

    # run_portfolio_optimizer()
