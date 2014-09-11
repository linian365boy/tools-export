package com.nian.export.demo;  

import java.io.IOException;
import java.io.OutputStream;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import com.nian.export.ColnameToField;
import com.nian.export.ExportExcelTool;
import com.nian.export.vo.QualityExportVo;

/** 
 * ClassName: QualityVoExportExcel
 * Desc: 质检导出
 * date: 2014-9-5 下午2:53:28
 * @author li.n1 
 * @since JDK 1.6 
 */
public class QualityVoExportExcel extends ExportExcelTool<QualityExportVo>{
	/**
	 * 适合导出到某个文件目录
	 * @param dataset
	 * @param path
	 * @param fileName
	 * @param title
	 */
	public QualityVoExportExcel(Collection<QualityExportVo> dataset,
			String path, String fileName, String title) {
		super(dataset, path);
		/* 设置Excel基本数据 */  
        // 定义列名与类属性名的对应关系  
        //  
        List<ColnameToField> colNameToField = new ArrayList<ColnameToField>() {  
            {  
                //add(new ColnameToField("orderNo", "订单号")); 如果注释，此列将不会导出
            	add(new ColnameToField("orderNo", "订单号"));
                add(new ColnameToField("expressCode", "收货快递单号"));  
                add(new ColnameToField("expressName", "快递公司"));  
                add(new ColnameToField("insideCode", "货品条码"));  
                add(new ColnameToField("supplierCode", "款色编码"));
                add(new ColnameToField("prodName", "商品名称"));  
                add(new ColnameToField("status", "质检状态"));  
                add(new ColnameToField("userName", "收货人姓名"));  
                add(new ColnameToField("saleType", "售后类型"));  
                add(new ColnameToField("remark", "退换货原因")); 
            }  
        };
        super.setTitle(title);// 标题，以查询的起止日期作为标题  
        super.setFileName(fileName);// Excel文件名(以站点名作为文件名)  
        super.setDateFormat("yyyy-MM-dd HH:mm:ss");  
        // 设置列名对应的属性名集合(其顺序与表头列名属性一致)  
        super.setColNameToField(colNameToField);  
	}
	
	/**
	 * web页面导出，弹出对话框
	 * @param dataset
	 * @param title
	 */
	public QualityVoExportExcel(Collection<QualityExportVo> dataset,String title) {
		super(dataset);
		/* 设置Excel基本数据 */  
        // 定义列名与类属性名的对应关系  
        List<ColnameToField> colNameToField = new ArrayList<ColnameToField>() {  
            {  
                //add(new ColnameToField("orderNo", "订单号")); 如果注释，此列将不会导出
            	add(new ColnameToField("orderNo", "订单号"));
                add(new ColnameToField("expressCode", "收货快递单号"));  
                add(new ColnameToField("expressName", "快递公司"));  
                add(new ColnameToField("insideCode", "货品条码"));  
                add(new ColnameToField("supplierCode", "款色编码"));
                add(new ColnameToField("prodName", "商品名称"));  
                add(new ColnameToField("status", "质检状态"));  
                add(new ColnameToField("userName", "收货人姓名"));  
                add(new ColnameToField("saleType", "售后类型"));  
                add(new ColnameToField("remark", "退换货原因")); 
            }  
        };
        super.setTitle(title);// 标题，以查询的起止日期作为标题  
        super.setDateFormat("yyyy-MM-dd HH:mm:ss");  
        // 设置列名对应的属性名集合(其顺序与表头列名属性一致)  
        super.setColNameToField(colNameToField);  
	}
	
	//controller实例
	/*
	@RequestMapping("/qualityExport")
    public @ResponseBody void qualityExport(ModelMap model, QualityQueryVo vo, Query query, 
    		HttpServletRequest request, HttpServletResponse response) {
        if (StringUtils.isBlank(vo.getQaTimeStart())) {
            vo.setQaTimeStart((new java.text.SimpleDateFormat("yyyy-MM-dd")).format(DateUtils.addMonths(new Date(), -3)));
        }
        if (StringUtils.isBlank(vo.getQaTimeEnd())) {
            vo.setQaTimeEnd((new java.text.SimpleDateFormat("yyyy-MM-dd")).format(new Date()));
        }
        vo.setMerchantCode(SessionUtil.getMerchantCodeFromSession(request));
        List<Map<String, Object>>  list = iQualityService.queryQualityAllListByVo(vo);
        List<QualityExportVo> exportList = prepareData(list);
        exportToExcel(exportList,response);
    }
    */
    
    /** 
	 * prepareCreate:命名文档跟导出 
	 * @author li.n1  
	 * @since JDK 1.6 
	 */ 
	/*
	private void exportToExcel(List<QualityExportVo> exportVoList, 
			HttpServletResponse response) {
		String title = "质检结果";
        OutputStream out = null;
        try {
        	QualityVoExportExcel excel = new QualityVoExportExcel(exportVoList, title);
        	response.setContentType("application/vnd.ms-excel");
        	response.setHeader("Content-Disposition", MessageFormat.format("attachment; filename={0}.xls", DateUtil.getCurrentDateTimeToStr()));
        	out = response.getOutputStream();
			excel.exportExcel(out);
		}catch(Exception e){
			logger.error("--质检结果导出发生异常！--");
		}finally {
			try {
				if(out!=null){
					out.close();
				}
			} catch (IOException e) {
				//e.printStackTrace();
				logger.info("IO异常，客户端窗口被关闭，导出失败！");
			}
		}
	}
	 */
	
	/** 
	 * prepareData:准备导出的数据 
	 * @author li.n1 
	 * @param list 
	 * @since JDK 1.6 
	 */ 
	/*
	private List<QualityExportVo> prepareData(List<Map<String, Object>> list) {
		List<AsmQcDetailVo> asmQcDetailVoList = null;
		List<QualityExportVo> exportVoList = new ArrayList<QualityExportVo>();
		QualityExportVo exportVo = null;
		List<SaleApplyBill> applyBills = null;
		Product product = null;
		SaleApplyBill bill = null;
		for(Map<String,Object> map : list){
			asmQcDetailVoList = (List<AsmQcDetailVo>) map.get("asmQcDetail");
			//根据订单号查询退换货申请单信息//有些订单没有退换货申请单信息
			applyBills = asmApiImpl.querySaleApplyBillListByOrderSubNo((String) map.get("order_sub_no"));
			for(AsmQcDetailVo asmQcDetailVo : asmQcDetailVoList){	//售后质检明细
				exportVo = new QualityExportVo();
				product = commodityBaseApiService.getProductByNo(asmQcDetailVo.getProdNo(), true);
				if(product!=null){
					exportVo.setSupplierCode(product.getCommodity().getSupplierCode());
				}
				exportVo.setExpressCode(asmQcDetailVo.getExpressCode());
				exportVo.setExpressName(asmQcDetailVo.getExpressName());
				exportVo.setInsideCode(asmQcDetailVo.getInsideCode());
				exportVo.setOrderNo((String) map.get("order_sub_no"));
				exportVo.setProdName(asmQcDetailVo.getProdName());
				exportVo.setUserName((String) map.get("user_name"));
				//已质检的信息，有申请单
					for (Iterator<SaleApplyBill> it = applyBills.iterator();it.hasNext();) {	
						bill = it.next();
						if(bill.getApplyNo().equals(asmQcDetailVo.getApplyNo())){
							exportVo.setRemark(bill.getRemark());
							exportVo.setSaleType("QUIT_GOODS".equals(bill.getSaleType()) ? "退货" : "换货");
							it.remove();
							break;
						}
					}
				if(null==exportVo.getSaleType()){
					exportVo.setSaleType(asmQcDetailVo.getQualityType());
				}
				exportVo.setStatus(asmQcDetailVo.getStatus());
				//exportVo.setSupplierCode(asmQcDetailVo.getSupplierCode());
				exportVoList.add(exportVo);
			}
		}
		return exportVoList;
	}
	*/
}
