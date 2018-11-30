import os
import pandas

mainpath = "/home/braytac/Documentos/Proveedores/"
file1 = "C_5817_28-11-2018 - Proveedor Suizo Arg.xls"
#file1 = "del sud - Proveedor Drog. del Sud.xlsx"
#file1 = "impositivo americana - Proveedor Monroe .xlsx"

fullpath = os.path.join(mainpath, file1)
dataframe = pandas.read_excel(fullpath)

jolistor_df = pandas.DataFrame({
                'Fecha Emisión':[],
                'Fecha Recepción':[],
                'Cpbte':[],
                'Tipo':[],
                'Suc.':[],
                'Número':[],
                'Razón Social/Denominación Proveedor':[],
                'Tipo Doc.':[],
                'CUIT':[],
                'Domicilio':[],
                'C.P.':[],
                'Pcia':[],
                'Cond Fisc':[],
                'Cód. Neto':[],
                'Neto Gravado':[],
                'Alíc.':[],
                'IVA Liquidado':[],
                'IVA Crédito':[],
                'Cód. NG/EX':[],
                'Conceptos NG/EX':[],
                'Cód. P/R':[],
                'Perc./Ret.':[],
                'Pcia P/R':[],
                'Total':[]
                })
                
tipo_suizo = pandas.Series(
                            ['sucursal', 'comprob', 'letra', 'terminal', 'numero', 'hojas',
                           'fecha', 'ctacte', 'clisuc', 'brutograv', 'brutonogra',
                           'descuento', 'netograv', 'ivari', 'percib', 'perciva', 'iint',
                           'netonograv', 'total', 'items', 'unidades', 'cai', 'vtocai',
                           'netoperf', 'fch_vto', 'concepto', 'xref_comp', 'xref_let',
                           'xref_ter', 'xref_num', 'cierre']
                            ).isin(dataframe.columns).all()



if tipo_suizo == True:
    dataframe.drop(['letra','terminal','hojas','ctacte','clisuc','brutonogra','descuento','netograv','ivari','percib','perciva','iint','netonograv','items','unidades','cai','vtocai','netoperf','fch_vto','xref_comp','xref_let','xref_ter','xref_num'], axis=1, inplace=True)
    #'Fecha Emisión ','Fecha Recepción','Cpbte','Tipo','Suc.','Número','Razón Social/Denominación Proveedor','Tipo Doc.','CUIT','Domicilio','C.P.','Pcia','Cond Fisc','Cód. Neto','Neto Gravado','Alíc.','IVA Liquidado','IVA Crédito','Cód. NG/EX','Conceptos NG/EX','Cód. P/R','Perc./Ret.','Pcia P/R','Total'
    col_dict_repl = {
                    'fecha':'Fecha Emisión',
                    'cierre':'Fecha Recepción',
                    'comprob':'Cpbte',
                    '':'Tipo',
                    'sucursal':'Suc.',
                    'numero':'Número',
                    '':'Razón Social/Denominación Proveedor',
                    '':'Tipo Doc.',
                    '':'CUIT',
                    '':'Domicilio',
                    '':'C.P.',
                    '':'Pcia',
                    '':'Cond Fisc',
                    '':'Cód. Neto',
                    'brutograv':'Neto Gravado',
                    '':'Alíc.',
                    '':'IVA Liquidado',
                    '':'IVA Crédito',
                    '':'Cód. NG/EX',
                    'concepto':'Conceptos NG/EX',
                    '':'Cód. P/R',
                    '':'Perc./Ret.',
                    '':'Pcia P/R',
                    'total':'Total' 
                    }   ## key→old name, value→new name
    dataframe.columns = [col_dict_repl.get(x, x) for x in dataframe.columns]
    #dataframe.rename(columns={'fecha':'Fecha Emisión'}, inplace=True)

    my_cols_list=[
                'Fecha Emisión',
                'Fecha Recepción',
                'Cpbte',
                'Tipo',
                'Suc.',
                'Número',
                'Razón Social/Denominación Proveedor',
                'Tipo Doc.',
                'CUIT',
                'Domicilio',
                'C.P.',
                'Pcia',
                'Cond Fisc',
                'Cód. Neto',
                'Neto Gravado',
                'Alíc.',
                'IVA Liquidado',
                'IVA Crédito',
                'Cód. NG/EX',
                'Conceptos NG/EX',
                'Cód. P/R',
                'Perc./Ret.',
                'Pcia P/R',
                'Total'
                ]

    dataframe = pandas.merge(dataframe, jolistor_df, how='outer') 
    dataframe = dataframe[ my_cols_list ]
    #dataframe.reindex(columns=[*dataframe.columns.tolist(), *my_cols_list], fill_value=True)
    # help(dataframe.reindex)
    dataframe.fillna("", inplace = True)
    dataframe.to_excel( mainpath+'/formateados_jolistor/'+file1+'_MOD.xls')
    
fullpath = ""
file2 = "impositivo americana - Proveedor Monroe .xlsx"
fullpath = os.path.join(mainpath, file2)
dataframe = pandas.read_excel(fullpath)

tipo_monroe = pandas.Series(
                            ['Tipo Linea', 'Tipo', 'Letra', 'Numero Formateado', 'Fecha',
                            'Caea Nro', 'Caea Vto', 'Cod Cliente', 'Cuit', 'Razon Social',
                            'Resumen', 'Tipo Pedido', 'Codigo Barra', 'Descripcion',
                            'Pcio Vta Pub', 'Base Excenta+Gravado', 'Iva', 'Otros Impuestos',
                            'Importe Total', 'Unidades', 'Lineas', 'Clasif Producto',
                            'Laboratorio', 'Condicion Iva', 'Porc Iva', 'Pcio Unitario']
                            ).isin(dataframe.columns).all()


if tipo_monroe == True:

    dataframe = dataframe.rename(columns=lambda x: x.strip().replace(' ','_'))
    #dataframe.query('Tipo_Linea != Cabecera')    
    dataframe = dataframe[dataframe.Tipo_Linea == 'Cabecera']
    
    col_dict_repl = {
                    'Fecha':'Fecha Emisión',
                    '':'Fecha Recepción',
                    '':'Cpbte',
                    'Tipo':'Tipo',
                    '':'Suc.',
                    'numero':'Número',
                    'Razon_Social':'Razón Social/Denominación Proveedor',
                    '':'Tipo Doc.',
                    'Cuit':'CUIT',
                    '':'Domicilio',
                    '':'C.P.',
                    '':'Pcia',
                    '':'Cond Fisc',
                    '':'Cód. Neto',
                    '':'Neto Gravado',
                    '':'Alíc.',
                    'Iva':'IVA Liquidado',
                    '':'IVA Crédito',
                    '':'Cód. NG/EX',
                    '':'Conceptos NG/EX',
                    '':'Cód. P/R',
                    '':'Perc./Ret.',
                    '':'Pcia P/R',
                    'Importe_Total':'Total' 
                    }   ## key→old name, value→new name
    dataframe.columns = [col_dict_repl.get(x, x) for x in dataframe.columns]    

    dataframe.drop(['Tipo_Linea',
                    'Letra',
                    'Numero_Formateado',
                    'Caea_Nro',
                    'Caea_Vto',
                    'Cod_Cliente',
                    'Resumen',
                    'Tipo_Pedido',
                    'Codigo_Barra',
                    'Descripcion',
                    'Pcio_Vta_Pub',
                    'Base_Excenta+Gravado',
                    'Otros_Impuestos',
                    'Unidades',
                    'Lineas',
                    'Clasif_Producto',
                    'Laboratorio',
                    'Condicion_Iva',
                    'Porc_Iva',
                    'Pcio_Unitario'], axis=1, inplace=True )

    dataframe = pandas.merge(dataframe, jolistor_df, how='outer') 
    dataframe = dataframe[ my_cols_list ]
    dataframe.fillna("", inplace = True)
    dataframe.to_excel( mainpath+'/formateados_jolistor/'+file2+'_MOD.xls')    

#dataframe	
