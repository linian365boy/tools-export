package com.nian.export;

/** 
 * ClassName:ColnameToField 
 * Date:     2014-9-5 下午1:56:58 
 * @author   li.n1 
 * @since    JDK 1.6 
 * @see      excel列跟属性对应关系
 */
public class ColnameToField {
		private String colname;//列名
		private String fieldName;//类属性名
		
		public ColnameToField(String fieldName, String colname){
			this.colname = colname;
			this.fieldName = fieldName;
		}

		public String getColname() {
			return colname;
		}

		public void setColname(String colname) {
			this.colname = colname;
		}

		public String getFieldName() {
			return fieldName;
		}

		public void setFieldName(String fieldName) {
			this.fieldName = fieldName;
		}
}
