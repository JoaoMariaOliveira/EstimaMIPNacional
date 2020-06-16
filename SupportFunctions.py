
import numpy as np
import pandas as pd
import string
import xlrd
import numpy.core.multiarray
from pandas import ExcelWriter
from pandas import ExcelFile

# ============================================================================================

def read_file_excel(sDirectory, sFileName, sSheetName):
    sFileName = sDirectory + sFileName
    mSheet = pd.read_excel(sFileName, sheet_name=sSheetName, header=None, index_col=None)
    return mSheet

# =============================================================
# Função que lê dados de um arquivo txt e retorna uma matriz
# lê de uma pasta input
# ============================================================================================

def read_file_txt(sFileName, sDirectory):
    sFileName = sDirectory + sFileName
    temp = np.loadtxt(sFileName, dtype=np.float64, skiprows= 1)
    return temp

# ============================================================================================
def load_intermediate_consumption(sDirectoryInput, sFileUses, sSheetIntermedConsum, nProduct , nSector ):
    mSheet = read_file_excel(sDirectoryInput, sFileUses, sSheetIntermedConsum)
    nColIni = 0
    nLinIni = 5
    vCodProduct = []
    vNameProduct = []
    vCodSector = []
    vNameSector = []
    mIntermConsum = np.zeros([nProduct+1, nSector+1], dtype=float)

    for p in range(nProduct):
        for s in range(nSector):
            mIntermConsum[p, s] = mSheet.values[nLinIni + p, nColIni + s + 2]
            mIntermConsum[nProduct, s] += mIntermConsum[p, s]

        mIntermConsum[p, nSector] = sum(mIntermConsum[p, :])

    mIntermConsum[nProduct, nSector] = sum(mIntermConsum[nProduct, :])
    for p in range(nProduct+1):
        vCodProduct.append(mSheet.values[nLinIni + p, nColIni])
        vNameProduct.append(mSheet.values[nLinIni + p, nColIni + 1])

    for s in range(nSector+1):
        sStr=mSheet.values[nLinIni - 2, nColIni + s + 2]
        vCodSector.append(sStr[0:4])
        vNameSector.append(sStr[5:])

    return vCodProduct, vNameProduct, vCodSector, vNameSector, mIntermConsum
# ============================================================================================


def load_demand(sDirectoryInput, sFileUses, sSheetDemand, nProduct, nColsDemand):
    mSheet = read_file_excel(sDirectoryInput, sFileUses, sSheetDemand)
    nColIni = 0
    nLinIni = 5
    mDemand = np.zeros([nProduct + 1, nColsDemand], dtype=float)
    vNameDemand = []
    for p in range(nProduct):
        for c in range (nColsDemand):
            mDemand [p, c]= mSheet.values[nLinIni + p, nColIni + c + 2]

    for c in range(nColsDemand):
        vNameDemand.append(mSheet.values[nLinIni - 2, nColIni + c + 2])
        mDemand [nProduct, c] =  sum(mDemand[:, c])

    return mDemand, vNameDemand
# ============================================================================================


def load_gross_added_value(sDirectoryInput, sFileUses, sSheetAddedValue, nSector, nRowsAV):
    mSheet = read_file_excel(sDirectoryInput, sFileUses, sSheetAddedValue)
    nColIni = 0
    nLinIni = 5
    vNameAddedValue = []
    mAddedValue = np.zeros([ nRowsAV, nSector + 1], dtype=float)
    for r in range(nRowsAV):
        vNameAddedValue.append(mSheet.values[nLinIni + r, nColIni])
        for s in range(nSector + 1):
            mAddedValue[r, s] = mSheet.values[nLinIni + r, nColIni + s + 1]

    return mAddedValue, vNameAddedValue
# ============================================================================================


def load_offer(sDirectoryInput, sFileResources, sSheetOffer, nProduct, nColsOffer):
    mSheet = read_file_excel(sDirectoryInput, sFileResources, sSheetOffer)
    nColIni = 0
    nLinIni = 5
    mOffer  = np.zeros([nProduct + 1, nColsOffer], dtype=float)
    vNameOffer    = []
    for p in range(nProduct):
        for c in range (nColsOffer):
            mOffer [p, c]= mSheet.values[nLinIni + p, nColIni + c + 2]

    for c in range(nColsOffer):
                vNameOffer.append(mSheet.values[nLinIni - 2, nColIni + c + 2])
                mOffer[nProduct, c] = sum(mOffer[:, c])

    return mOffer, vNameOffer
# ============================================================================================


def load_production(sDirectoryInput, sFileResources, sSheetProduction, nProduct, nSector):
    mSheet = read_file_excel(sDirectoryInput, sFileResources, sSheetProduction,)
    nColIni = 0
    nLinIni = 5
    mProduction = np.zeros([nProduct + 1, nSector + 1], dtype=float)

    for p in range(nProduct):
        for s in range(nSector):
            mProduction[p, s] = mSheet.values[nLinIni + p, nColIni + s + 2]
            mProduction[nProduct, s] += mProduction[p, s]

        mProduction[p, nSector] = sum(mProduction[p, :])

    mProduction[nProduct, nSector] = sum(mProduction[nProduct, :])
    return mProduction
# ============================================================================================

def load_import(sDirectoryInput, sFileResources, sSheetImport, nProduct):
    mSheet = read_file_excel(sDirectoryInput, sFileResources, sSheetImport)
    nColIni = 0
    nLinIni = 5
    vImport = np.zeros([nProduct+1, 1], dtype=float)
    for p in range(nProduct):
        vImport[p, 0] = mSheet.values[nLinIni + p, nColIni + 2]

    vImport[nProduct, 0] = sum(vImport[:, 0])
    return vImport
# ============================================================================================

def calculation_distribution_matrix_nat(mIntermConsum, mDemand):
    mTotalConsum =np.concatenate((mIntermConsum, mDemand), axis=1)
    nRow, nCol =  np.shape(mTotalConsum)
    mDistribution = np.zeros([nRow, nCol], dtype=float)
    for r in range(nRow):
        for c in range(nCol):
            if (mTotalConsum [r, nCol-1]==0):
                mDistribution[r, c] = 0.0
            else:
                mDistribution[r, c] = mTotalConsum[r, c] / mTotalConsum [r, nCol-1]



    return mDistribution
# ============================================================================================

def calculation_margin_nat(mAlpha, mDemand, nColRef, vRowErase):
    nAux = 0
    for r in range(vRowErase[0],vRowErase[1]+1):
        nAux += mDemand[r, nColRef]

    nRow, nCol = np.shape(mAlpha)
    mMatrixOutput = np.zeros([nRow, nCol], dtype=float)

    for c in range(nCol):
        nTot = 0
        nTotMargin = 0
        for r in range(nRow-1):
            mMatrixOutput[r, c] = mAlpha[r, c] * mDemand[r, nColRef]
            if (np.isnan(mMatrixOutput[r, c])):
                mMatrixOutput[r, c] = 0
            if r < vRowErase[0] or r > vRowErase[1]:
                nTot += mMatrixOutput[r, c]

        for r in range(vRowErase[0], vRowErase[1]+1):
            nMultip = mDemand[r, nColRef] / nAux
            if (np.isnan(nMultip)):
               nMultip = 0.0

            mMatrixOutput[r, c] = nTot * (-1.) * nMultip
            nTotMargin += mMatrixOutput[r, c]

        mMatrixOutput[nRow-1, c] = nTot + nTotMargin

    return mMatrixOutput
# ============================================================================================

def calculation_margin_reg(mAlpha, mDemand, nColRef, vRowErase, nStates, nProducts):
    mAux = np.zeros([nStates, 1], dtype=float)
    mTotMargin  = np.zeros([nStates, 1], dtype=float)
#    nAux=0
    for s in range(nStates):
        for r in range(s*nProducts + vRowErase[0], s*nProducts + vRowErase[1]+1):
           mAux[s, 0] += mDemand[r, nColRef]
#           nAux += mDemand[r, nColRef]

    nRow, nCol = np.shape(mAlpha)
    mMatrixOutput = np.zeros([nRow, nCol], dtype=float)
    for s in range(nStates):
        for c in range(nCol):
            nTot = 0
#            nTotMargin = 0
            for r in range(s*nProducts, (s+1)*nProducts):
                mMatrixOutput[r, c] = mAlpha[r, c] * mDemand[r, nColRef]
                if (np.isnan(mMatrixOutput[r, c])):
                    mMatrixOutput[r, c] = 0
                if r < (s*nProducts + vRowErase[0]) or r > (s*nProducts + vRowErase[1]):
                    nTot += mMatrixOutput[r, c]

            for r in range(s*nProducts + vRowErase[0], s*nProducts + vRowErase[1]+1):
                nMultip = mDemand[r, nColRef] / mAux[s,0]
#                nMultip = mDemand[r, nColRef] / nAux
                if (np.isnan(nMultip)):
                    nMultip = 0.0

                mMatrixOutput[r, c] = nTot * (-1.) * nMultip
 #               nTotMargin += mMatrixOutput[r, c]
                mTotMargin [s,0] += mMatrixOutput[r, c]

#            mMatrixOutput[nRow-1, c] = nTot + nTotMargin
    ar1 = sum(mTotMargin[:, 0])
    ar2 = sum(mAux[:, 0])
    return mMatrixOutput
# ============================================================================================

def calculation_internal_matrix_nat(mAlpha, mDemand, nColRef):
    nRow, nCol = np.shape(mAlpha)
    mMatrixOutput = np.zeros([nRow, nCol], dtype=float)
    for c in range(nCol):
        nAux=0
        for r in range(nRow-1):
            mMatrixOutput[r, c] = mAlpha[r, c] * mDemand[r, nColRef]
            if (np.isnan(mMatrixOutput[r, c])):
                mMatrixOutput[r, c] = 0
            nAux += mMatrixOutput[r, c]

        mMatrixOutput[nRow-1, c] = nAux

    return mMatrixOutput
# ============================================================================================

def calculation_internal_matrix_reg(mAlpha, mDemand, nColRef):
    nRow, nCol = np.shape(mAlpha)
    mMatrixOutput = np.zeros([nRow, nCol], dtype=float)
    for c in range(nCol):
 #       nAux=0
        for r in range(nRow-1):
            mMatrixOutput[r, c] = mAlpha[r, c] * mDemand[r, nColRef]
            if (np.isnan(mMatrixOutput[r, c])):
                mMatrixOutput[r, c] = 0
  #          nAux += mMatrixOutput[r, c]

#        mMatrixOutput[nRow-1, c] = nAux

    return mMatrixOutput

# ============================================================================================

def GDP_Calculation (mMIPGeral, nSector):
    vNameGDP  =[]
    vNameColGDP = []
    vNameColGDP.append('Valores')
    vGDP =  np.zeros([17], dtype=float)
    nRowMIP, nColMIP = np.shape(mMIPGeral)

    vNameGDP.append('PIB pela ótica do produto')
    vNameGDP.append('Produção')
    vNameGDP.append('Impostos')
    vNameGDP.append('Consumo Intermediário')

    nProducao = mMIPGeral[nSector, nColMIP - 1]
    nImpostos = mMIPGeral[nSector + 2, nColMIP - 1] + mMIPGeral[nSector + 3, nColMIP - 1] + \
                mMIPGeral[nSector + 4, nColMIP - 1] + mMIPGeral[nSector + 5, nColMIP - 1]
    nCI = mMIPGeral[nSector + 6, nSector]
    nPIBProduto = nProducao + nImpostos - nCI

    vGDP[0] = nPIBProduto
    vGDP[1] = nProducao
    vGDP[2] = nImpostos
    vGDP[3] = nCI

    vNameGDP.append('PIB pela ótica da Renda')
    vNameGDP.append('Remuneração dos empregados')
    vNameGDP.append('Rendimento Misto Bruto')
    vNameGDP.append('EOB')
    vNameGDP.append('Impostos líquidos sobre a produção e importação')

    nRemunEmpregados = mMIPGeral[nSector + 8, nSector]
    nEOB = mMIPGeral[nSector + 16, nSector]
    nRendiMistoBruto = mMIPGeral[nSector + 15, nSector]
    nImpLiqProdImport = nImpostos + mMIPGeral[nSector + 17, nSector] + mMIPGeral[nSector + 18, nSector]
    nPIBrenda = nRemunEmpregados + nEOB + nRendiMistoBruto + nImpLiqProdImport


    vGDP[4] = nPIBrenda
    vGDP[5] = nRemunEmpregados
    vGDP[6] = nRendiMistoBruto
    vGDP[7] = nEOB
    vGDP[8] = nImpLiqProdImport

    vNameGDP.append('PIB pela ótica da despesa')
    vNameGDP.append('Consumo das Famílias')
    vNameGDP.append('Consumo do Governo')
    vNameGDP.append('Consumo das ISFLSF')
    vNameGDP.append('FBCF')
    vNameGDP.append('Variação do estoque')
    vNameGDP.append('exportação de bens e serviços')
    vNameGDP.append('importação de bens e serviços (-)')

    nExportTotal = mMIPGeral[nSector + 6, nSector + 1]
    nGovernConsum = mMIPGeral[nSector + 6, nSector + 2]
    nISFLSFConsum = mMIPGeral[nSector + 6, nSector + 3]
    nFamilyConsum = mMIPGeral[nSector + 6, nSector + 4]
    nFBCF = mMIPGeral[nSector + 6, nSector + 5]
    nColEstockVar = mMIPGeral[nSector + 6, nSector + 6]
    nimporTotal = mMIPGeral[nSector + 1, nColMIP - 1]
    nPIBDespesa = nExportTotal + nGovernConsum + nISFLSFConsum + nFamilyConsum + nFBCF + nColEstockVar - nimporTotal

    vGDP[9] = nPIBDespesa
    vGDP[10] = nFamilyConsum
    vGDP[11] = nGovernConsum
    vGDP[12] = nISFLSFConsum
    vGDP[13] = nFBCF
    vGDP[14] = nColEstockVar
    vGDP[15] = nimporTotal
    vGDP[16] = nExportTotal

    return vGDP, vNameGDP, vNameColGDP

# ============================================================================================
# Função que grava dados em um arquivo excel
# Grava em uma pasta output e pode gravar várias planilhas em um mesmo arquivo
# ============================================================================================

def write_data_excel(sDirectoryOutput, sFileName, lSheetName, lDataSheet, lRowsLabel, lColsLabel, vUseHeader):
    Writer = pd.ExcelWriter(sDirectoryOutput+sFileName, engine='xlsxwriter')
    df=[]
    for each in range(len(lSheetName)):
        if vUseHeader[each] == True:
            df.append(pd.DataFrame(lDataSheet[each],  index=lRowsLabel[each], columns=lColsLabel[each], dtype=float))
            df[each].to_excel(Writer, lSheetName[each], header=True, index=True)
        else:
            df.append(pd.DataFrame(lDataSheet[each],  dtype=float))
            df[each].to_excel(Writer, lSheetName[each], header=False, index=False)

    Writer.save()

# ============================================================================================
# Função que grava dados em um arquivo excel
# Grava em uma pasta output e pode gravar uma só planilha em um mesmo arquivo
# ============================================================================================

def write_file_excel(sDirectoryOutput, sFileName, sSheet, mData, vRows, vCols, lUseHeader):
    Writer = pd.ExcelWriter(sDirectoryOutput+sFileName, engine='xlsxwriter')
    if lUseHeader == True:
        df=pd.DataFrame(mData, index=vRows, columns=vCols, dtype=float)
        df.to_excel(Writer, sheet_name=sSheet, header=True, index=True)
    else:
        df=pd.DataFrame(mData, dtype=float)
        df.to_excel(Writer, sheet_name=sSheet, header=False, index=False)
    Writer.save()

# ============================================================================================