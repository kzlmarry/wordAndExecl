package com.cplatform.app.util;

import com.cplatform.app.domain.*;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import tk.mybatis.mapper.util.StringUtil;

import javax.servlet.ServletOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.TreeMap;


/**
 * excel工具类,支持批量导出
 * @author lizewu
 *
 */
public class ExportExcel {
    
    /**
     * 将卡证申请信息导入到excel文件中去
     * @param workList 工作列表
     * @param out 输出表
     */
    public static void exportWorkExcel(List<CardApplication> workList,ServletOutputStream out)
    {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        try{
            //1.创建工作簿
            HSSFWorkbook workbook = new HSSFWorkbook();
            //1.1创建合并单元格对象
			CellRangeAddress calltips = new CellRangeAddress(0,0,0,4);//起始行,结束行,起始列,结束列
            //1.2头标题样式
            HSSFCellStyle headStyle = createCellStyle(workbook,(short)16);
            //1.3列标题样式
            HSSFCellStyle colStyle = createCellStyle(workbook,(short)13);
            //2.创建工作表
            HSSFSheet sheet = workbook.createSheet("卡证Excel表");
            //2.1加载合并单元格对象
			sheet.addMergedRegion(calltips);
            //设置默认列宽
            sheet.setDefaultColumnWidth(25);
            //3.创建行
			//第一行内容
			HSSFRow headRow = sheet.createRow(0);
			HSSFCell headCell = headRow.createCell(0);
			headCell.setCellStyle(headStyle);
			headCell.setCellValue("卡证申请表");
            
            //3.2创建表说明行;并且设置头标题
//            HSSFRow row2 = sheet.createRow(2);
//            if(userCard != null)
//            {
//				HSSFCell cell3 = row2.createCell(5);
//				Date createTime = userCard.getMealCreateTime();
//				String newTime = sdf.format(createTime);
//				//加载单元格样式
//				cell3.setCellValue(newTime);
//			}

            //3.3创建列标题;并且设置列标题
            HSSFRow row3 = sheet.createRow(1);
            String[] titles = {"序号","申请人","申请时间","申请内容","结果"};
            for(int i=0;i<titles.length;i++)
            {
                HSSFCell cell2 = row3.createCell(i);
                //加载单元格样式
                cell2.setCellStyle(colStyle);
                cell2.setCellValue(titles[i]);
            }
            
            
            //4.操作单元格;将用户列表写入excel
            if(workList != null)
            {
                for(int j=0;j<workList.size();j++)
                {
                    //创建数据行,前面有二行,头标题行，表说明行和列标题行
                    HSSFRow row4 = sheet.createRow(j+2);
                    HSSFCell cell1 = row4.createCell(0);
                    cell1.setCellValue(j+1);
                    HSSFCell cell2 = row4.createCell(1);
                    cell2.setCellValue(workList.get(j).getName());
                    HSSFCell cell3 = row4.createCell(2);
                    //dateFormat_yyyyMMdd.format(card.getUpdateTime())
                    Date deliveryTime = workList.get(j).getCreateTime();
                    cell3.setCellValue(sdf.format(deliveryTime));
                    HSSFCell cell4 = row4.createCell(3);
                    String type = null;
                    String status = null;
                    if(1 == workList.get(j).getType()) {
                    	type = "餐卡、门禁卡办理";
                    }else if(2 == workList.get(j).getType()) {
                    	type = "车辆通行证办理";
                    }
                    cell4.setCellValue(type);
                    if(0 == workList.get(j).getStatus()) {
                    	status = "驳回";
                    }else if(1 == workList.get(j).getStatus()) {
                    	status = "待审批";
                    }else if(2 == workList.get(j).getStatus()) {
                    	status = "办理中";
                    }else if(3 == workList.get(j).getStatus()) {
                    	status = "完成";
                    }
                    HSSFCell cell5 = row4.createCell(4);
                    cell5.setCellValue(status);
                }
            }
            
            //5.输出
            workbook.write(out);
            workbook.close();
        }catch(Exception e)
        {
            e.printStackTrace();
        }
    }
    /**
     * 将卡证申请信息导入到excel文件中去
     * @param workList 工作列表
     * @param out 输出表
     */
    public static void exportCancelExcel(List<CardCancelApplication> workList,ServletOutputStream out)
    {
    	SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
    	try{
    		//1.创建工作簿
    		HSSFWorkbook workbook = new HSSFWorkbook();
    		//1.1创建合并单元格对象
    		CellRangeAddress calltips = new CellRangeAddress(0,0,0,4);//起始行,结束行,起始列,结束列
    		//1.2头标题样式
    		HSSFCellStyle headStyle = createCellStyle(workbook,(short)16);
    		//1.3列标题样式
    		HSSFCellStyle colStyle = createCellStyle(workbook,(short)13);
    		//2.创建工作表
    		HSSFSheet sheet = workbook.createSheet("卡证Excel表");
    		//2.1加载合并单元格对象
    		sheet.addMergedRegion(calltips);
    		//设置默认列宽
    		sheet.setDefaultColumnWidth(25);
    		//3.创建行
    		//第一行内容
    		HSSFRow headRow = sheet.createRow(0);
    		HSSFCell headCell = headRow.createCell(0);
    		headCell.setCellStyle(headStyle);
    		headCell.setCellValue("卡证注销申请表");
    		
    		//3.2创建表说明行;并且设置头标题
//            HSSFRow row2 = sheet.createRow(2);
//            if(userCard != null)
//            {
//				HSSFCell cell3 = row2.createCell(5);
//				Date createTime = userCard.getMealCreateTime();
//				String newTime = sdf.format(createTime);
//				//加载单元格样式
//				cell3.setCellValue(newTime);
//			}
    		
    		//3.3创建列标题;并且设置列标题
    		HSSFRow row3 = sheet.createRow(1);
    		String[] titles = {"序号","申请人","申请时间","申请内容","结果"};
    		for(int i=0;i<titles.length;i++)
    		{
    			HSSFCell cell2 = row3.createCell(i);
    			//加载单元格样式
    			cell2.setCellStyle(colStyle);
    			cell2.setCellValue(titles[i]);
    		}
    		
    		
    		//4.操作单元格;将用户列表写入excel
    		if(workList != null)
    		{
    			for(int j=0;j<workList.size();j++)
    			{
    				//创建数据行,前面有二行,头标题行，表说明行和列标题行
    				HSSFRow row4 = sheet.createRow(j+2);
    				HSSFCell cell1 = row4.createCell(0);
    				cell1.setCellValue(j+1);
    				HSSFCell cell2 = row4.createCell(1);
    				cell2.setCellValue(workList.get(j).getName());
    				HSSFCell cell3 = row4.createCell(2);
    				//dateFormat_yyyyMMdd.format(card.getUpdateTime())
    				Date deliveryTime = workList.get(j).getCreateTime();
    				cell3.setCellValue(sdf.format(deliveryTime));
    				HSSFCell cell4 = row4.createCell(3);
    				String type = null;
    				String status = null;
    				if(0 == workList.get(j).getType()) {
    					type = "餐卡、门禁卡注销";
    				}else if(1 == workList.get(j).getType()) {
    					type = "车辆通行证注销";
    				}
    				cell4.setCellValue(type);
    				if(0 == workList.get(j).getStatus()) {
    					status = "驳回";
    				}else if(1 == workList.get(j).getStatus()) {
    					status = "待审批";
    				}else if(2 == workList.get(j).getStatus()) {
    					status = "办理中";
    				}else if(3 == workList.get(j).getStatus()) {
    					status = "完成";
    				}
    				HSSFCell cell5 = row4.createCell(4);
    				cell5.setCellValue(status);
    			}
    		}
    		
    		//5.输出
    		workbook.write(out);
    		workbook.close();
    	}catch(Exception e)
    	{
    		e.printStackTrace();
    	}
    }
    /**
     * 将卡证申请信息导入到excel文件中去
     * @param workList 工作列表
     * @param out 输出表
     */
    public static void exportMealExcel(List<MealApplication> workList,ServletOutputStream out)
    {
    	SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
    	try{
    		//1.创建工作簿
    		HSSFWorkbook workbook = new HSSFWorkbook();
    		//1.1创建合并单元格对象
    		CellRangeAddress calltips = new CellRangeAddress(0,0,0,4);//起始行,结束行,起始列,结束列
    		//1.2头标题样式
    		HSSFCellStyle headStyle = createCellStyle(workbook,(short)16);
    		//1.3列标题样式
    		HSSFCellStyle colStyle = createCellStyle(workbook,(short)13);
    		//2.创建工作表
    		HSSFSheet sheet = workbook.createSheet("卡证Excel表");
    		//2.1加载合并单元格对象
    		sheet.addMergedRegion(calltips);
    		//设置默认列宽
    		sheet.setDefaultColumnWidth(25);
    		//3.创建行
    		//第一行内容
    		HSSFRow headRow = sheet.createRow(0);
    		HSSFCell headCell = headRow.createCell(0);
    		headCell.setCellStyle(headStyle);
    		headCell.setCellValue("餐卡购买表");
    		
    		//3.2创建表说明行;并且设置头标题
//            HSSFRow row2 = sheet.createRow(2);
//            if(userCard != null)
//            {
//				HSSFCell cell3 = row2.createCell(5);
//				Date createTime = userCard.getMealCreateTime();
//				String newTime = sdf.format(createTime);
//				//加载单元格样式
//				cell3.setCellValue(newTime);
//			}
    		
    		//3.3创建列标题;并且设置列标题
    		HSSFRow row3 = sheet.createRow(1);
    		String[] titles = {"序号","申请人","申请时间","申请内容","结果"};
    		for(int i=0;i<titles.length;i++)
    		{
    			HSSFCell cell2 = row3.createCell(i);
    			//加载单元格样式
    			cell2.setCellStyle(colStyle);
    			cell2.setCellValue(titles[i]);
    		}
    		
    		
    		//4.操作单元格;将用户列表写入excel
    		if(workList != null)
    		{
    			for(int j=0;j<workList.size();j++)
    			{
    				//创建数据行,前面有二行,头标题行，表说明行和列标题行
    				HSSFRow row4 = sheet.createRow(j+2);
    				HSSFCell cell1 = row4.createCell(0);
    				cell1.setCellValue(j+1);
    				HSSFCell cell2 = row4.createCell(1);
    				cell2.setCellValue(workList.get(j).getName());
    				HSSFCell cell3 = row4.createCell(2);
    				//dateFormat_yyyyMMdd.format(card.getUpdateTime())
    				Date deliveryTime = workList.get(j).getCreateTime();
    				cell3.setCellValue(sdf.format(deliveryTime));
    				HSSFCell cell4 = row4.createCell(3);
    				String type = "餐劵购买";
    				String status = null;
    				cell4.setCellValue(type);
    				if(0 == workList.get(j).getStatus()) {
    					status = "驳回";
    				}else if(1 == workList.get(j).getStatus()) {
    					status = "待审批";
    				}else if(2 == workList.get(j).getStatus()) {
    					status = "办理中";
    				}else if(3 == workList.get(j).getStatus()) {
    					status = "完成";
    				}
    				HSSFCell cell5 = row4.createCell(4);
    				cell5.setCellValue(status);
    			}
    		}
    		
    		//5.输出
    		workbook.write(out);
    		workbook.close();
    	}catch(Exception e)
    	{
    		e.printStackTrace();
    	}
    }
    /**
     * 
     * @param workbook
     * @param fontsize
     * @return 单元格样式
     */
    @SuppressWarnings("deprecation")
	private static HSSFCellStyle createCellStyle(HSSFWorkbook workbook, short fontsize) {
        HSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//水平居中
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直居中
        //创建字体
        HSSFFont font = workbook.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font.setFontHeightInPoints(fontsize);
        //加载字体
        style.setFont(font);
        return style;
    }
    
}