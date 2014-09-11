package com.nian.export;

import org.junit.Test;

/**
 * Unit test for simple ExportExcelTest.
 */
//@ContextConfiguration(locations = { "classpath:applicationContext-Test.xml" })
public class ExportExcelTest {
	@Test
	public void testQualityExport(){
		/*//Query query = new Query(30);
		QualityQueryVo vo = new QualityQueryVo();
		vo.setMerchantCode("SP20130821678648");
		vo.setQaTimeStart("2014-06-09");
		vo.setQaTimeEnd("2014-09-09");
		List<Map<String, Object>> list = qualityService.queryQualityAllListByVo(vo);
		System.out.println(list.size());
		assertTrue(list.size() >0);
		
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
				exportVo.setSupplierCode(product.getCommodity().getSupplierCode());
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
		Calendar now = Calendar.getInstance();
		String title = "质检结果";
        String fileName = "质检结果"+now.get(Calendar.YEAR)+
        		"-"+now.get(Calendar.MONTH)+"-"+
        		now.get(Calendar.DATE)+"-"+now.get(Calendar.MILLISECOND)+".xls";
        String path = "D:\\";  
        QualityVoExportExcel excel = new QualityVoExportExcel(exportVoList, path, fileName, title);  
        excel.exportExcel(); */ 
	}
}
