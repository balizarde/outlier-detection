# ---------------------------------------------------------
# Outlier detection by commute-time and Euclidean distances
# 
# Sercan Taha Ahi, Nov 2011 (tahaahi at gmail dot com)
# ---------------------------------------------------------
from numpy import *
import sys
import math
import scipy as Sci
import scipy.linalg
import string
import xlrd
import matplotlib.pyplot as plt
from pylab import *

def parse_stats(fname):  
    wb = xlrd.open_workbook(fname)
    sh = wb.sheet_by_index(0)
    
    PLY = sh.col_values(0)
    PLY = PLY[1:len(PLY)]
    n_players = len(PLY)
    print "\tn_players = " + str(n_players)
    
    TEAM = sh.col_values(0)
    TEAM = TEAM[1:n_players]
    
    POS = zeros((n_players,1)); AGE = zeros((n_players,1)); G = zeros((n_players,1))
    MIN = zeros((n_players,1)); MINPG = zeros((n_players,1))
    PTS = zeros((n_players,1)); PTSPG = zeros((n_players,1))
    FGM = zeros((n_players,1)); FGA = zeros((n_players,1)); FGP = zeros((n_players,1))
    FTM = zeros((n_players,1)); FTA = zeros((n_players,1)); FTP = zeros((n_players,1))
    TPM = zeros((n_players,1)); TPA = zeros((n_players,1)); TPP = zeros((n_players,1))
    ORB = zeros((n_players,1)); ORBPG = zeros((n_players,1))
    DRB = zeros((n_players,1)); DRBPG = zeros((n_players,1))
    TRB = zeros((n_players,1)); TRBPG = zeros((n_players,1))
    AST = zeros((n_players,1)); ASTPG = zeros((n_players,1))
    STL = zeros((n_players,1)); STLPG = zeros((n_players,1))
    BLK = zeros((n_players,1)); BLKPG = zeros((n_players,1))
    TO  = zeros((n_players,1)); TOPG  = zeros((n_players,1))
    PF  = zeros((n_players,1)); PFPG  = zeros((n_players,1))
    for i in range(0, n_players):
        if sh.cell(i+1,2).value == 'C':
            POS[i,0] = 1
        elif sh.cell(i+1,2).value == 'F':
            POS[i,0] = 2
        elif sh.cell(i+1,2).value == 'G':
            POS[i,0] = 3
        else:
            print "Error: Unknown position"
        
        AGE[i,0]   = sh.cell(i+1,6).value
        G[i,0]     = sh.cell(i+1,7).value
        MIN[i,0]   = sh.cell(i+1,8).value
        MINPG[i,0] = sh.cell(i+1,9).value
        PTS[i,0]   = sh.cell(i+1,10).value
        PTSPG[i,0] = sh.cell(i+1,11).value
        FGM[i,0]   = sh.cell(i+1,12).value
        FGA[i,0]   = sh.cell(i+1,13).value
        FGP[i,0]   = sh.cell(i+1,14).value
        FTM[i,0]   = sh.cell(i+1,15).value
        FTA[i,0]   = sh.cell(i+1,16).value
        FTP[i,0]   = sh.cell(i+1,17).value
        TPM[i,0]   = sh.cell(i+1,18).value
        TPA[i,0]   = sh.cell(i+1,19).value
        TPP[i,0]   = sh.cell(i+1,20).value
        ORB[i,0]   = sh.cell(i+1,21).value
        ORBPG[i,0] = sh.cell(i+1,22).value
        DRB[i,0]   = sh.cell(i+1,23).value
        DRBPG[i,0] = sh.cell(i+1,24).value
        TRB[i,0]   = sh.cell(i+1,25).value
        TRBPG[i,0] = sh.cell(i+1,26).value
        AST[i,0]   = sh.cell(i+1,27).value
        ASTPG[i,0] = sh.cell(i+1,28).value
        STL[i,0]   = sh.cell(i+1,29).value
        STLPG[i,0] = sh.cell(i+1,30).value
        BLK[i,0]   = sh.cell(i+1,31).value
        BLKPG[i,0] = sh.cell(i+1,32).value
        TO[i,0]    = sh.cell(i+1,33).value
        TOPG[i,0]  = sh.cell(i+1,34).value
        PF[i,0]    = sh.cell(i+1,35).value
        PFPG[i,0]  = sh.cell(i+1,36).value
        
	X = concatenate((POS, AGE, G, MINPG, PTSPG), axis=1)
    X = concatenate((X, FGA/G, FGM/G), axis=1)
    X = concatenate((X, FTA/G, FTM/G), axis=1)
    X = concatenate((X, TPA/G, TPM/G), axis=1)
    X = concatenate((X, ORBPG, DRBPG, ASTPG, STLPG), axis=1)
    X = concatenate((X, BLKPG, TOPG, PFPG), axis=1)
    
    #X = [POS AGE G MINPG PPG FGAPG FGMPG FTAPG FTMPG TPAPG TPMPG ...
    #     ORBPG DRBPG ASTPG STLPG BLKPG TOPG PFPG];
    
    #X = concatenate((POS, PTSPG, TRBPG, ASTPG, STLPG, BLKPG), axis=1)
    
    print "\tX: " + str(X.shape[0]) + "x" + str(X.shape[1])
    
    return X, PLY

    
def normalize(X):
    n_samples = X.shape[0]
    n_features = X.shape[1]
    print "\tn_samples  = " + str(n_samples)
    print "\tn_features = " + str(n_features)
    
    Xmean = X.mean(axis=0)
    Xstd  = X.std(axis=0)
    
    Y = zeros((n_samples, n_features))
    for i in range(0,n_samples):
        Y[i,:] = X[i,:] - Xmean
        Y[i,:] = Y[i,:] / Xstd
        
    #for i in range(0,n_features):
    #    print "\t\t" + str(Y[:,i].std())
    #print "\tmax = " + str(Y.max())
    #print "\tmin = " + str(Y.min())
    return Y

    
def dist_mahalanobis(X):
    n_samples = X.shape[0]
    n_features = X.shape[1]
    
    Xmean = X.mean(axis=0)
    Xc = zeros((n_samples, n_features))
    for i in range(0,n_samples):
        Xc[i,:] = X[i,:] - Xmean
    
    S = dot(Xc.T, Xc) / (n_features-1)
    Si = linalg.inv(S)
    
    D = zeros((n_samples, n_samples))
    for i in range(0,n_samples):
        a = X[i,:]
        a.reshape(1,n_features)
        for j in range(0,n_samples):
            b = X[j,:]
            b.reshape(1,n_features)
            D[i,j] = dot(dot(a-b,Si),(a-b).T)
    
    return D

    
def dist_euclidean(X):
    n_samples = X.shape[0]
    #n_features = X.shape[1]
    
    D = zeros((n_samples, n_samples))
    for i in range(0,n_samples):
        a = X[i,:]
        for j in range(0,n_samples):
            b = X[j,:]
            D[i,j] = linalg.norm(a-b)
    
    return D

    
def dist_commute_time(X):
    n_samples = X.shape[0]
    #n_features = X.shape[1]
    
    E = dist_euclidean(X)
    Estd = E.std()
    A = exp(-E**2 / Estd**2);
    D = zeros((n_samples,n_samples))
    for i in range(0,n_samples):
        A[i,i] = 0
        D[i,i] = A[i,:].sum(dtype=float)
    
    V = A.sum(dtype=float)
    print "\tGraph volume = " + str(V)
    
    #D = diag(A.sum(axis=1));
    L = D - A;
    Lp = linalg.pinv(L);
    
    CTD = zeros((n_samples,n_samples))
    for i in range(0,n_samples):
        for j in range(0,n_samples):
            CTD[i,j] = V * (Lp[i,i] + Lp[j,j] - 2*Lp[i,j])
    
    #CTD = CTD / CTD.max()
    #E = E / E.max()
    
    return CTD, E
    
    
def get_top_n_outliers(D, knn, n, dtype):
    n_samples = D.shape[0]
    
    Dtop = zeros((n_samples))
    for i in range(0,n_samples):
        idx = argsort(D[i,:])
        idx = idx[1:knn+1]
        Dtop[i] = D[i,idx].mean()
    
    idx = argsort(-Dtop)
    idx = idx[0:n]
    
    fname = dtype + ".png"
    x = arange(n_samples)
    y = sort(Dtop)
    plt.figure()
    plt.plot(x, y, 'b-')
    plt.grid()
    plt.title(dtype + " Distance")
    plt.xlabel("Player Index")
    plt.ylabel("Distance")
    plt.draw()
    plt.savefig(fname, dpi=300)
    
    return idx
    
    
def main():
    if len(sys.argv) == 1:
        target_year = 1997
    else:
        target_year = int(sys.argv[1])
    fname = str(target_year) + '-' + str(target_year+1) + '.xls'
    
    X, PLY = parse_stats(fname)
    X = normalize(X)
    CTD, E = dist_commute_time(X)
    M = dist_mahalanobis(X)
    
    knn = 10
    n = 10
    
    print "\nOutliers based on commute time distance:"
    dtype = "Commute Time"
    idx = get_top_n_outliers(CTD, knn, n, dtype)
    for i in range(0,n):
        print str(i+1).rjust(3) + "\t" + PLY[idx[i]]
    
    print "\nOutliers based on Euclidean distance:"
    dtype = "Euclidean"
    idx = get_top_n_outliers(E, knn, n, dtype)
    for i in range(0,n):
        print str(i+1).rjust(3) + "\t" + PLY[idx[i]]
    
    print "\nOutliers based on Mahalanobis distance:"
    dtype = "Mahalanobis"
    idx = get_top_n_outliers(M, knn, n, dtype)
    for i in range(0,n):
        print str(i+1).rjust(3) + "\t" + PLY[idx[i]]
    
main()
