import re

from datetime import datetime, timedelta

import string

from lxml import etree, objectify


class CT_Chart_Container(objectify.ObjectifiedElement):
    """
    ``<p:graphicFrame>`` element, which is a container for a table, a chart,
    or another graphical object.
    """
    DATATYPE_TABLE = 'http://schemas.openxmlformats.org/drawingml/2006/table'
    
    _chart_tmpl = (
      
      '<c:chartSpace %s>'
      '<c:date1904 val="0"/>'
      '<c:lang val="en-US"/>'
      '<c:roundedCorners val="0"/>'
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
      '<mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">'
      '<c14:style val="102"/>'
      '</mc:Choice>'
      '<mc:Fallback>'
      '<c:style val="2"/>'
      '</mc:Fallback>'
      '</mc:AlternateContent>'
      '<c:chart>'
      '<c:autoTitleDeleted val="0"/>'
      '<c:plotArea>'
      '<c:layout/>'
      
      '</c:plotArea>'    
      '<c:legend>'
      '<c:legendPos val="r"/>'
          '<c:layout/>'
          '<c:overlay val="0"/>'
        '</c:legend>'
        '<c:plotVisOnly val="1"/>'
        '<c:dispBlanksAs val="gap"/>'
        '<c:showDLblsOverMax val="0"/>'
      '</c:chart>'
      '<c:txPr>'
      '<a:bodyPr/>'
      '<a:lstStyle/>'
      '<a:p>'
          '<a:pPr>'
            '<a:defRPr sz="1800"/>'
          '</a:pPr>'
          '<a:endParaRPr lang="en-US"/>'
        '</a:p>'
      '</c:txPr>'
      '<c:externalData r:id="rId1">'
        '<c:autoUpdate val="0"/>'
      '</c:externalData>'
    '</c:chartSpace>'

       % nsdecls('c','a','r')
)
    
    _chart_tmpl1 = (
      
      
        '<c:chart %s>'
        '<c:autoTitleDeleted val="0"/>'
        '<c:plotArea>'
        '<c:lineChart>'
           
        '</c:lineChart>' 
        '</c:plotArea>'
        '</c:chart>' % nsdecls('c')
    )


    @staticmethod
    def new_chart_wrapper():

        xml = CT_Chart_Container._chart_tmpl 
        graphicFrame = oxml_fromstring(xml)

        objectify.deannotate(graphicFrame, cleanup_namespaces=True)
        return graphicFrame

    @staticmethod
    def new_chart(data,headings_xlsx):

        headings=[[]]
        for i,e in enumerate(headings_xlsx):
            headings[i].append(headings_xlsx[i])
            if i == len(headings_xlsx)-1:
                continue
            else:
                headings.append([])

        
        chartFrame =  CT_Chart_Container.new_chart_wrapper()

        # set type of contained graphic to table
        chartData = chartFrame['chart'].plotArea  

        # add chart data element tree
        chart = CT_ChartCell.new_chart(data,headings,headings_xlsx)
        chartData.append(chart)
        chart =CT_plotArea.new_plot_vars()
        chartData.append(chart)
        
        chart =CT_plotArea.plot_val_vars()
        chartData.append(chart)
        objectify.deannotate(chartData, cleanup_namespaces=True)
        chartXML = etree.tostring(chartFrame, pretty_print=True,xml_declaration = True, encoding='UTF-8', standalone="yes")
        return chartXML
    
class CT_ChartCell(objectify.ObjectifiedElement):
    
    _tbl_tmpl = (
        '<dummy>'
        '</dummy>'
    )
    
    @staticmethod
    def new_chart(data,headings,headings_xlsx):
 
        
        #xml = new('c:lineChart')
        #grouping = sub_elm(xml,'c:grouping',val="standard")
        #varycolors = sub_elm(xml,'c:varyColors',val="0")
        
        xml = new('c:lineChart')
        grouping = sub_elm(xml,'c:grouping',val="standard")
        varycolors = sub_elm(xml,'c:varyColors',val="0")
        
        
        chart = xml

        
        for idx, val in enumerate(headings):
            
            if idx == 0:
                continue
            else:
                ref_series = "Sheet1!$"+ string.uppercase[idx]+ '$1' 
                
                ref_cat    = 'Sheet1!$A$2:$A$'+ str(len(data[1])+1)
                
                ref_val   = "Sheet1!$"+ string.uppercase[idx]+ '$2' + ":$"  + string.uppercase[idx] + "$"+str(len(data[0])+1) 
    
                tr = sub_elm(chart, 'c:ser')
                tx = sub_elm(tr,'c:idx',val=str(idx-1))
                tx = sub_elm(tr,'c:order',val=str(idx-1))
                
                # Category for Legend
                
                tx = sub_elm(tr,'c:tx')
                ostrRef = CT_ChartCell.strRef_node(val,ref_series)
                tx.append(ostrRef) 
                
                marker = sub_elm(tr,'c:marker')
                sym    = sub_elm (marker,'c:symbol', val = "none")
        
                
                # X-axis Categories
                
                cat = sub_elm(tr,'c:cat')
                
                ostrRef = CT_ChartCell.strRef_node(data[0],ref_cat)
                cat.append(ostrRef) 
                
                
                # Series Values - Y-Axis
                
                val = sub_elm(tr,'c:val')
                
                ostrRef = CT_ChartCell.numRef_node(data[idx],ref_val)
                val.append(ostrRef) 
                
                marker = sub_elm(tr,'c:smooth', val = "0")
        
            # c:dLbls Settings for Line Chart 
                
        odLbls = sub_elm(chart,'c:dLbls')
        legend = sub_elm(odLbls,'c:showLegendKey',val ="0")
        legend = sub_elm(odLbls,'c:showVal',val ="0")
        legend = sub_elm(odLbls,'c:showCatName',val ="0")
        legend = sub_elm(odLbls,'c:showSerName',val ="0")
        legend = sub_elm(odLbls,'c:showPercent',val ="0")
        legend = sub_elm(odLbls,'c:showBubbleSize',val ="0")
            
        mark   = sub_elm(chart,'c:marker', val="1")
        mark   =  sub_elm(chart,'c:smooth', val="0")
        mark   =  sub_elm(chart,'c:axId', val="172167552")
        mark   =  sub_elm(chart,'c:axId', val="172169088")
  
   

            
        # add specified number of rows and columns

        
        objectify.deannotate(chart, cleanup_namespaces=True)
        return chart
    
    
    @staticmethod
    def strRef_node(series,series_ref):
        """ create new strRef node c:strCache -> <c:ptCount> -> <c:pt idx =0> -> <c:v> 
        """
        #xml = CT_ChartCell._tbl_tmpl 
        #series= ['Series 1','Series 2','Series 3']
        #series_ref = 'Sheet1!$B$1'
        data = new('c:strRef')
        strRef= data
        f = sub_elm(strRef,'c:f')
        ostrCache = sub_elm(strRef,'c:strCache')
        optCount = sub_elm(ostrCache,'c:ptCount',val=str(len(series)))
        
        for idx, val in enumerate(series):
            opt = sub_elm(ostrCache,'c:pt',idx=str(idx))
            #ov = sub_elm(opt,'c:v')
            opt.v= val
        
        
        strRef.f = series_ref
       
        objectify.deannotate(data, cleanup_namespaces=True)
        return data
    @staticmethod
    def numRef_node(series,series_ref):
        """ create new  c:strCache -> <c:ptCount> -> <c:pt idx =0> -> <c:v> 
        """
        #xml = CT_ChartCell._tbl_tmpl 
        #series= ['Series 1','Series 2','Series 3']
        #series_ref = 'Sheet1!$B$1'
        data = new('c:numRef')
        strRef= data
        f = sub_elm(strRef,'c:f')
        ostrCache = sub_elm(strRef,'c:numCache')
        ostrCache.formatCode = 'General'
        optCount = sub_elm(ostrCache,'c:ptCount',val=str(len(series)))
        
        for idx, val in enumerate(series):
            opt = sub_elm(ostrCache,'c:pt',idx=str(idx))
            #ov = sub_elm(opt,'c:v')
            opt.v= val
        
        
        strRef.f = series_ref
       
        objectify.deannotate(data, cleanup_namespaces=True)
        return data

class CT_plotArea(objectify.ObjectifiedElement):
    """ <c:catAx>
        <c:axId val="172167552"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="172169088"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="172169088"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="172167552"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>"""

    @staticmethod
    def new_plot_vars():
        ocatAx = new('c:catAx')
        ocatAx_l= sub_elm(ocatAx,'c:axId', val ="172167552")
        ocatAx_l= sub_elm(ocatAx,'c:scaling')
        ocatAx_l= sub_elm(ocatAx_l,'c:orientation', val ="minMax")
        ocatAx_l= sub_elm(ocatAx,'c:delete',val="0")
        ocatAx_l= sub_elm(ocatAx,'c:axPos',val="b")
        ocatAx_l= sub_elm(ocatAx,'c:majorTickMark',val="out")
        ocatAx_l= sub_elm(ocatAx,'c:minorTickMark',val="none")
        ocatAx_l= sub_elm(ocatAx,'c:tickLblPos',val="nextTo")
        ocatAx_l= sub_elm(ocatAx,'c:crossAx',val="172169088")
        ocatAx_l= sub_elm(ocatAx,'c:crosses',val="autoZero")
        ocatAx_l= sub_elm(ocatAx,'c:auto',val="1")
        ocatAx_l= sub_elm(ocatAx,'c:lblAlgn',val="ctr")
        ocatAx_l= sub_elm(ocatAx,'c:lblOffset',val="100")
        ocatAx_l= sub_elm(ocatAx,'c:noMultiLvlLbl',val="0")
        
  
        objectify.deannotate(ocatAx, cleanup_namespaces=True)
        return ocatAx  
        
    @staticmethod
    def plot_val_vars():
     
        ovalAx = new('c:valAx')
        ovalAx_l= sub_elm(ovalAx,'c:axId', val ="172169088")
        ovalAx_l= sub_elm(ovalAx,'c:scaling')
        ovalAx_l= sub_elm(ovalAx_l,'c:orientation', val ="minMax")
        ovalAx_l= sub_elm(ovalAx,'c:delete',val="0")
        ovalAx_l= sub_elm(ovalAx,'c:axPos',val="l")
        ovalAx_l= sub_elm(ovalAx,'c:majorGridlines')
        ovalAx_l= sub_elm(ovalAx,'c:numFmt',formatCode="General" ,sourceLinked ="1")
        ovalAx_l= sub_elm(ovalAx,'c:majorTickMark',val="out")
        ovalAx_l= sub_elm(ovalAx,'c:minorTickMark',val="none")
        ovalAx_l= sub_elm(ovalAx,'c:tickLblPos',val="nextTo")
        ovalAx_l= sub_elm(ovalAx,'c:crossAx',val="172167552")
        ovalAx_l= sub_elm(ovalAx,'c:crosses',val="autoZero")
        ovalAx_l= sub_elm(ovalAx,'c:crossBetween',val="between")
        
      
        objectify.deannotate(ovalAx, cleanup_namespaces=True)
        return ovalAx  
        


        

