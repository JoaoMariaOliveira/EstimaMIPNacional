# ============================================================================================
# Main Module of Building Input output National
# Coded to Python by João Maria de Oliveira -
# ============================================================================================
import yaml
import numpy as np
import SupportFunctions as Support
import sys
#from multiprocessing import Pool, cpu_count
import time



# ============================================================================================
# read Parameters into Confg file
# ============================================================================================
conf = yaml.load(open('config.yaml', 'r'), Loader=yaml.FullLoader)

sDirectoryInput  = conf['sDirectoryInput']
sDirectoryOutput = conf['sDirectoryOutput']

sSheetIntermedConsum = conf['sSheetIntermedConsum']
sSheetDemand = conf['sSheetDemand']
sSheetAddedValue = conf['sSheetAddedValue']
sSheetOffer = conf['sSheetOffer']
sSheetProduction = conf['sSheetProduction']
sSheetImport = conf['sSheetImport']

nColsDemand = conf['nColsDemand']
nColsDemandEach = conf['nColsDemandEach']
nColsOffer = conf['nColsOffer']
nRowsAV = conf['nRowsAV']
nColExport = conf['nColExport']
nColISFLSFConsum =  conf['nColISFLSFConsum']
nColGovConsum = conf['nColGovConsum']
nColFBCF = conf['nColFBCF']
nColStockVar = conf['nColStockVar']
nColMarginTrade = conf['nColMarginTrade']
nColMarginTransport = conf['nColMarginTransport']
nColIPI = conf['nColIPI']
nColICMS = conf['nColICMS']
nColOtherTaxes = conf['nColOtherTaxes']
nColImport = conf['nColImport']
nColImportTax = conf['nColImportTax']

nDimension = conf['nDimension']
nYear = conf['nYear']
lAdjustMargins = conf['lAdjustMargins']
mAdjust = conf['mAdjust']
vProducts = conf['vProducts']
vSectors = conf['vSectors']
vRowsTrade = conf['vRowsTrade']
vRowsTransp = conf['vRowsTransp']
vColsTrade = conf['vColsTrade']
vColsTransp = conf['vColsTransp']

# nProducts - Número de produtos de acordo com a dimensão da MIP
# nSectors - Número de atividades de acordo com a dimensão da MIP
# lAdjustMargins - True se ajusta as margens de comércio e transporte para apenas um produto e uma atividade
nProducts = vProducts [nDimension]
nSectors = vSectors [nDimension]
vRowsTradeElim  = vRowsTrade[nDimension]
vRowsTranspElim = vRowsTransp[nDimension]
vColsTradeElim  = vColsTrade[nDimension]
vColsTranspElim = vColsTransp[nDimension]

if lAdjustMargins:
   sAdjustMargins = '_Agreg'
else:
    sAdjustMargins = ''

# nRowTotalProduction - Número da linha do total da produção
# nRowAddedValueGross - Número da Linha do valor adicionado Bruto
# nColTotalDemand - Numero da coluna da demanda total na  tabela demanda
# nColFinalDemand - Numero da coluna da demanda total na  tabela demanda
nRowTotalProduction = nRowsAV - 2
nRowAddedValueGross = 0
nColTotalDemand = nColsDemand - 1
nColFinalDemand = nColsDemand - 2

# sFileUses             - Arquivo de usos - Demanda
# sFileResources        - Arquivo de recursos
# sFileSheet - nome do arquivo de saida contendo as = tabelas
sFileUses      = str(nSectors)+'_tab2_'+str(nYear)+'.xls'
sFileResources = str(nSectors)+'_tab1_'+str(nYear)+'.xls'
sFileSheetNat = 'MIP_Nat_'+str(nYear)+'_'+str(nSectors)+sAdjustMargins+'.xlsx'
sFileNameOutput = str(nYear)+'_'+str(nSectors)+sAdjustMargins+'.xlsx'


if __name__ == '__main__':

    nBeginModel = time.perf_counter()
    sTimeBeginModel = time.localtime()
    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print("Running Model by year ",nYear," for ",nProducts,"products x ", nSectors, "sectors ")
    print("Begin at ", time.strftime("%d/%b/%Y - %H:%M:%S",sTimeBeginModel ))
    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    # ============================================================================================
    # Import values from TRUs
    # ============================================================================================
    vCodProduct, vNameProduct, vCodSector, vNameSector, mIntermConsumNat = Support.load_intermediate_consumption \
        (sDirectoryInput, sFileUses, sSheetIntermedConsum, nProducts, nSectors)

    mDemandNat, vNameDemand = Support.load_demand(sDirectoryInput, sFileUses, sSheetDemand, nProducts, nColsDemand)

    mAddedValueNat, vNameAddedValue = Support.load_gross_added_value \
        (sDirectoryInput, sFileUses, sSheetAddedValue, nSectors, nRowsAV)

    mOfferNat, vNameOffer = Support.load_offer(sDirectoryInput, sFileResources, sSheetOffer, nProducts, nColsOffer)

    mProductionNat = Support.load_production(sDirectoryInput, sFileResources, sSheetProduction, nProducts, nSectors)

    vImportNat = Support.load_import(sDirectoryInput, sFileResources, sSheetImport, nProducts)

    # ============================================================================================
    # Adjusting Trade and Transport for Products and for Sectors
    #  ============================================================================================
    sTimeIntermediate = time.localtime()
    print(time.strftime("%d/%b/%Y - %H:%M:%S", sTimeIntermediate), " - Adjusting trade and transports for products/sectors")
    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    lAdjust = lAdjustMargins
    nAdjust = 0
    while lAdjust:
        if (mAdjust[nAdjust]==1):
            if nAdjust == 0:
                nRowIni = vRowsTradeElim[0]
                nRowFim = vRowsTradeElim[1]
                vNameProduct[nRowIni] = 'Comércio'
                nColIni = vColsTradeElim[0]
                nColFim = vColsTradeElim[1]
                vNameSector[nColIni] = 'Comércio'
            else:
                nRowIni = vRowsTranspElim[0]
                nRowFim = vRowsTranspElim[1]
                vNameProduct[nRowIni] = 'Transporte'
                nColIni = vColsTranspElim[0]
                nColFim = vColsTranspElim[1]
                vNameSector[nColIni] = 'Transporte'

            for nElim in range(nRowIni + 1, nRowFim + 1):
                vNameProduct[nElim] = 'x'
                vImportNat[nRowIni] += vImportNat[nElim]
                vImportNat[nElim] = 0.0

            for i in range(nColsOffer):
                for nElim in range(nRowIni + 1, nRowFim + 1):
                    mOfferNat[nRowIni, i] += mOfferNat[nElim, i]
                    mOfferNat[nElim, i] = 0.0

            for i in range(nSectors + 1):
                for nElim in range(nRowIni + 1, nRowFim + 1):
                    mProductionNat[nRowIni, i] += mProductionNat[nElim, i]
                    mProductionNat[nElim, i] = 0.0
                    mIntermConsumNat[nRowIni, i] += mIntermConsumNat[nElim, i]
                    mIntermConsumNat[nElim, i] = 0.0

            for i in range(nColsDemand):
                for nElim in range(nRowIni + 1, nRowFim + 1):
                    mDemandNat[nRowIni, i] += mDemandNat[nElim, i]
                    mDemandNat[nElim, i] = 0.0

            for nElim in range(nColIni + 1, nColFim + 1):
                vNameSector[nElim] = 'x'

            for i in range(nRowsAV):
                for nElim in range(nColIni + 1, nColFim + 1):
                    mAddedValueNat[i, nColIni] += mAddedValueNat[i, nElim]
                    mAddedValueNat[i, nElim] = 0.0

            for i in range(nProducts + 1):
                for nElim in range(nColIni + 1, nColFim + 1):
                    mProductionNat[i, nColIni] += mProductionNat[i, nElim]
                    mProductionNat[i, nElim] = 0.0
                    mIntermConsumNat[i, nColIni] += mIntermConsumNat[i, nElim]
                    mIntermConsumNat[i, nElim] = 0.0

        nAdjust += 1
        if nAdjust == 2:
            lAdjust = False

    # ============================================================================================
    # Calculating totals of Matrix's
    # ============================================================================================

    # ============================================================================================
    # Calculating Coeficients without Stock Variation
    #  ============================================================================================
    sTimeIntermediate = time.localtime()
    print(time.strftime("%d/%b/%Y - %H:%M:%S",sTimeIntermediate)," - Calculating Coeficients without Stocks")
    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")

    mDemandNatWithoutStock = np.copy(mDemandNat)
    # zering Stock variation
    mDemandNatWithoutStock[:, nColStockVar] = 0.0
    # Recalculanting  final and total demand
    for p in range(nProducts + 1):
        mDemandNatWithoutStock[p, nColTotalDemand] = mDemandNat[p, nColTotalDemand] - mDemandNat[p, nColStockVar]
        mDemandNatWithoutStock[p, nColFinalDemand] = mDemandNat[p, nColFinalDemand] - mDemandNat[p, nColStockVar]

    mDistributionNat = Support.calculation_distribution_matrix_nat(mIntermConsumNat, mDemandNatWithoutStock)

    # ============================================================================================
    # Calculating Arrays internally distributed by alphas
    #  ============================================================================================
    nColMarginTrade = 1
    mMarginTradeNat = Support.calculation_margin_nat(mDistributionNat, mOfferNat, nColMarginTrade, vRowsTradeElim)

    nColMarginTransport = 2
    mMarginTransportNat = Support.calculation_margin_nat(mDistributionNat, mOfferNat, nColMarginTransport, vRowsTranspElim)

    nColIPI = 4
    mIPINat = Support.calculation_internal_matrix_nat(mDistributionNat, mOfferNat, nColIPI)

    nColICMS = 5
    mICMSNat = Support.calculation_internal_matrix_nat(mDistributionNat, mOfferNat, nColICMS)

    nColOtherTaxes = 6
    mOtherTaxesNat = Support.calculation_internal_matrix_nat(mDistributionNat, mOfferNat, nColOtherTaxes)

    # ============================================================================================
    # Calculating Coeficients without exports and Stock Variation
    #  ============================================================================================
    sTimeIntermediate = time.localtime()
    print(time.strftime("%d/%b/%Y - %H:%M:%S",sTimeIntermediate), " - Calculating Coeficients without Stocks and exports")
    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")

    mDemandNatWithoutExport = np.copy(mDemandNatWithoutStock)

    # zering exports
    mDemandNatWithoutExport[:, nColExport] = 0

    # Recalculanting  final and total demand
    for p in range(nProducts + 1):
        mDemandNatWithoutExport[p, nColTotalDemand] = mDemandNatWithoutStock[p, nColTotalDemand] - \
                                                      mDemandNatWithoutStock[p, nColExport]
        mDemandNatWithoutExport[p, nColFinalDemand] = mDemandNatWithoutStock[p, nColFinalDemand] - \
                                                      mDemandNatWithoutStock[p, nColExport]


    mDistributionNatWithoutExport = Support.calculation_distribution_matrix_nat(mIntermConsumNat, mDemandNatWithoutExport)

    # ============================================================================================
    # Calculating Arrays internally distributed by alphas without exports
    #  ============================================================================================
    mImportNat = Support.calculation_internal_matrix_nat(mDistributionNatWithoutExport, vImportNat, nColImport)
    mImportTaxNat = Support.calculation_internal_matrix_nat(mDistributionNatWithoutExport, mOfferNat, nColImportTax)

    # ============================================================================================
    # Calculating the Matrix of Consum with base Price -
    #  ============================================================================================
    sTimeIntermediate = time.localtime()
    print(time.strftime("%d/%b/%Y - %H:%M:%S",sTimeIntermediate), " - Calculating BasePrices uses matrix")
    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")

    mTotalConsumNat = np.concatenate((mIntermConsumNat, mDemandNat), axis=1)
    mConsumBasePriceNat = mTotalConsumNat - mMarginTradeNat - mMarginTransportNat - mIPINat - mICMSNat - mOtherTaxesNat\
                          - mImportNat - mImportTaxNat

    sTimeIntermediate = time.localtime()
    print(time.strftime("%d/%b/%Y - %H:%M:%S",sTimeIntermediate), " - Creating U, E, D, B, Y, A and Z Matrix")
    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    # Calculating U, E and Q Matrix's - National and Regional
    mUNat = mConsumBasePriceNat[0:nProducts, 0:nSectors]
    mENat = mConsumBasePriceNat[0:nProducts, nSectors+1:nSectors + nColsDemandEach + 1]
    vQNat = np.zeros([nProducts], dtype=float)
    for p in range(nProducts):
        vQNat[p] = sum(mUNat[p,:]) + sum(mENat[p,:])

    nQNat = sum(vQNat[:])
    # Creating V, Ql and D Matrix's  - National and Regional
    mVNat = mProductionNat[0:nProducts, 0:nSectors].T
    vQlNat = np.zeros([nProducts], dtype=float)
    for p in range(nProducts):
        vQlNat[p] = sum(mVNat[:, p])

    nQlNat = sum(vQlNat[:])
    if (nQNat != nQlNat):
        print("error: difference between Q= ",nQNat," and Q'= ",nQlNat)

    mDNat = np.zeros([nSectors, nProducts], dtype=float)

    # Alternative to calc mDNat
    # mDProv = np.dot(mVNat , np.diagflat(1/vQlNat) )

    for s in range(nSectors):
        for p in range(nProducts):
            if (vQlNat[p]==0):
                mDNat[s, p] = 0
            else:
                mDNat[s, p] = mVNat[s, p] / vQlNat[p]

    # Calculating X and Bn Matrix's - National and Regional
    vXNat = mAddedValueNat[nRowTotalProduction, 0:nSectors]
    mBnNat = np.zeros([nProducts, nSectors], dtype=float)
    for r in range(nProducts):
        for c in range(nSectors):
            if vXNat[c] == 0:
                mBnNat[r, c] = 0.0
            else:
                mBnNat[r, c] = mUNat[r, c] / vXNat[c]

    # Calculating Y, I, A  and Z Matrix's  - National and Regional
    mYNat = (np.dot(mDNat, mENat))
    mINat = np.eye(nSectors)
    mANat = np.dot(mDNat, mBnNat)
    mZNat = np.zeros([nSectors, nSectors], dtype=float)

    for r in range(nSectors):
        for c in range(nSectors):
            mZNat[r, c] = mANat[r, c] * vXNat[c]

    # Calculating Leontief Matrix  ( Sector x Sector )
    mLeontiefNat =  np.linalg.inv(mINat - mANat)

    mMIPNat = np.concatenate((mZNat, mYNat), axis=1)

    sTimeIntermediate = time.localtime()
    print(time.strftime("%d/%b/%Y - %H:%M:%S",sTimeIntermediate), " - Creating Bm, M, T, W and I-O Matrix")
    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")

    # Calculating Bm  - National and Regional
    mBmNat = np.zeros([nProducts, nSectors], dtype=float)
    for r in range(nProducts):
        for c in range(nSectors):
            if vXNat[c] == 0:
                mBmNat[r, c] = 0.0
            else:
                mBmNat[r, c] = mImportNat[r, c] / vXNat[c]

    # Calculating MIntCons, MDemand, TIntCons and TDemand Matrix's  - National and Regional
    vMIntConsNat = np.zeros([1, nSectors], dtype=float)
    vMDemandNat = np.zeros([1, nColsDemandEach], dtype=float)
    for s in range(nSectors):
        vMIntConsNat[0,s] = sum(mImportNat[0:nProducts, s])

    for s in range(nColsDemandEach):
        vMDemandNat[0,s] = sum(mImportNat[0:nProducts, nSectors+1+s])

    mTIntConsNat = np.zeros([4, nSectors], dtype=float)
    mTDemandNat = np.zeros([4, nColsDemandEach], dtype=float)
    for s in range(nSectors):
        mTIntConsNat[0,s] = sum(mImportTaxNat[0:nProducts, s])
        mTIntConsNat[1,s] = sum(mIPINat[0:nProducts, s])
        mTIntConsNat[2,s] = sum(mICMSNat[0:nProducts, s])
        mTIntConsNat[3,s] = sum(mOtherTaxesNat[0:nProducts, s])

    for s in range(nColsDemandEach):
        mTDemandNat[0, s] = sum(mImportTaxNat[0:nProducts, nSectors+1+s])
        mTDemandNat[1, s] = sum(mIPINat[0:nProducts, nSectors+1+s])
        mTDemandNat[2, s] = sum(mICMSNat[0:nProducts, nSectors+1+s])
        mTDemandNat[3, s] = sum(mOtherTaxesNat[0:nProducts, nSectors+1+s])


    # Calculating WIntCons  - National and Regional
    vWIntConsNat = np.zeros([1, nSectors], dtype=float)
    vWDemandNat = np.zeros([1, nColsDemandEach], dtype=float)
    vTotCINat =  np.zeros([1, nSectors], dtype=float)
    for s in range(nSectors):
        vWIntConsNat[0, s] = mAddedValueNat[nRowAddedValueGross, s]

    vMNat = np.concatenate((vMIntConsNat, vMDemandNat), axis=1)
    mTNat = np.concatenate((mTIntConsNat, mTDemandNat), axis=1)
    vWNat = np.concatenate((vWIntConsNat, vWDemandNat), axis=1)

    mMIPGeralNat = np.concatenate((mMIPNat, vMNat, mTNat, vWNat), axis=0)

    nRow, nCol = np.shape(mMIPGeralNat)
    vTotRowNat = np.zeros([nRow, 1], dtype=float)
    vTotColNat = np.zeros([1, nCol], dtype=float)

    for r in range(nRow):
        vTotRowNat[r, 0] = np.sum(mMIPGeralNat[r, :])

    for c in range(nCol):
        vTotColNat[0, c] = np.sum(mMIPGeralNat[:, c])

    vNameMTWX = [['Importação'], ['II'], ['IPI'], ['ICMS'], ['OILL'], ['VA'] ]

    sTimeIntermediate = time.localtime()
    print(time.strftime("%d/%b/%Y - %H:%M:%S", sTimeIntermediate), " - Writing National Mip Matrix")
    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")

    vDataSheet = []
    vSheetName = []
    vRowsLabel = []
    vColsLabel = []
    vUseHeader = []

#    vGDP, vNameGDP, vNameColGDP = Support.GDP_Calculation(mMIPGeral, nSectors)

    vDataSheet.append(mProductionNat)
    vSheetName.append('CI_Recursos_Input')
    vRowsLabel.append(vNameProduct)
    vColsLabel.append(vNameSector)
    vUseHeader.append(False)


    vDataSheet.append(mIntermConsumNat)
    vSheetName.append('CI_Usos_Input')
    vRowsLabel.append(vNameProduct)
    vColsLabel.append(vNameSector)
    vUseHeader.append(True)





    vDataSheet.append(vImportNat)
    vSheetName.append('Importação_Input')
    vRowsLabel.append(vNameProduct)
    vColsLabel.append(['Importação'])
    vUseHeader.append(True)

    vNameRowsMipGeralNat= vNameSector[0:nSectors]+vNameMTWX
    vNameColsMipGeralNat= vNameSector[0:nSectors]+vNameDemand[0:6]


    vDataSheet.append(mMIPGeralNat)
    vSheetName.append('MIP_Nat')
    vRowsLabel.append(vNameRowsMipGeralNat)
    vColsLabel.append(vNameColsMipGeralNat)
    vUseHeader.append(True)

    Support.write_data_excel(sDirectoryOutput, sFileSheetNat, vSheetName, vDataSheet, vRowsLabel, vColsLabel, vUseHeader)

    print("Terminou ")
    sys.exit(0)


