package com.cplatform.operator.auth.controller;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;

import com.common.page.Page;
import com.common.util.WordUtil;
import com.common.web.base.BaseController;
import com.cplatform.app.domain.CardApplication;
import com.cplatform.app.domain.CardApproval;
import com.cplatform.app.domain.CardCancelApplication;
import com.cplatform.app.domain.CardTransaction;
import com.cplatform.app.domain.MealApplication;
import com.cplatform.app.domain.MealApproval;
import com.cplatform.app.domain.MealTransaction;
import com.cplatform.app.domain.MealVoucher;
import com.cplatform.app.domain.TuserCard;
import com.cplatform.app.mapper.MealApplicationMapper;
import com.cplatform.app.mapper.MealApprovalMapper;
import com.cplatform.app.mapper.MealTransactionMapper;
import com.cplatform.app.service.MealApplicationService;
import com.cplatform.app.service.MealApprovalService;
import com.cplatform.app.util.ExportExcel;
import com.cplatform.app.util.ImgUtils;
import com.github.pagehelper.PageHelper;
import com.github.pagehelper.PageInfo;

import tk.mybatis.mapper.common.Mapper;
import tk.mybatis.mapper.entity.Example;
import tk.mybatis.mapper.entity.Example.Criteria;

/**
 * 餐卡购买
 * @author ZhouLe
 *
 */
@Controller
@RequestMapping("meal")
public class MealController extends BaseController<MealApplication, MealApplication> {

	@Autowired
	private MealApplicationMapper mapper;
	@Autowired
	private MealApprovalMapper approvalMapper;
	@Autowired
	private MealTransactionMapper transactionMapper;
	@Autowired
	private MealApplicationService service;
	
	
	@Override
	public Mapper<MealApplication> getMapper() {
		return mapper;
	}
	/**
     * 餐卷购买
     * @param approval
     * @param page
     * @return
     */
	@RequestMapping("list")
	public ModelAndView list(MealApplication application, Page page) {
		ModelAndView mv = new ModelAndView("meal/list");
		PageHelper.startPage(page.getPageNo(), page.getPageSize());
		List<MealApplication> list = service.FindAll(application);
		PageInfo<MealApplication> pageList = new PageInfo<>(list);
    	mv.addObject("page",pageList);
        mv.addObject("cur_nav", "meal");
    	return mv;
	}
	
	/**
     * 详细数据
     * @param approval
     * @param page
     * @return
     */
	@RequestMapping("selectOne")
	public ModelAndView SelectOne(Long id) {
		ModelAndView mv = new ModelAndView("meal/selectOne");
		//查询申请表信息
		MealApplication application = mapper.selectByPrimaryKey(id);
		//查询审批表信息
		Example example = new Example(MealApproval.class);
		Criteria criteria = example.createCriteria();
		criteria.andEqualTo("mealApplicationId",id);
		List<MealApproval> approvalList = approvalMapper.selectByExample(example);
		//查询办理表信息
		Example ex = new Example(MealTransaction.class);
		Criteria cr = example.createCriteria();
		cr.andEqualTo("mealApplicationId",id);
		List<MealTransaction> transactionList = transactionMapper.selectByExample(ex);
		if(approvalList.size() > 0) {
			mv.addObject("approvalList",approvalList.get(0));
		}
    	mv.addObject("cardApproval",application);
    	if(transactionList.size() > 0) {
    		mv.addObject("transactionList",transactionList.get(0));
    	}
        mv.addObject("cur_nav", "selectOne");
    	return mv;
	}
	
	/**
	 * 餐劵购买申请信息导出excel
	 */
	@RequestMapping("exportMealExcel")
	public void exportMealExcel(MealApplication application,HttpServletRequest request,HttpServletResponse response)throws Exception {
		List<MealApplication> list = service.FindAll(application);
		//2.导出
		//这里设置的文件格式是application/x-excel
		response.setContentType("application/x-excel");
		String fileName ="餐劵购买申请信息.xls";
		response.setHeader("Content-Disposition", "attachment;filename=" + new String(fileName.getBytes(), "ISO-8859-1"));
		ServletOutputStream outputStream = response.getOutputStream();
		ExportExcel.exportMealExcel(list, outputStream);
		if(outputStream != null) {
			outputStream.close();
		}
	}
	
	/***
	 * 导出详细数据
	 * 
	 * @return
	 * @throws Exception
	 */
	@RequestMapping(value = "export", method = RequestMethod.GET)
	@ResponseBody
	public void export(HttpServletRequest request, HttpServletResponse response, Long id) throws Exception {
		String realPath = request.getSession().getServletContext().getRealPath("/");
		String path = realPath + "/tmp/zip/";
		String excelPath = realPath + "/tmp/excel/";
		String dirOrFile = realPath + "/tmp/excel/";
		//查询申请表信息
		MealApplication application = mapper.selectByPrimaryKey(id);
		//查询审批表信息
		Example example = new Example(MealApproval.class);
		Criteria criteria = example.createCriteria();
		criteria.andEqualTo("mealApplicationId",id);
		List<MealApproval> approvalList = approvalMapper.selectByExample(example);
		//查询办理表信息
		Example ex = new Example(MealTransaction.class);
		Criteria cr = example.createCriteria();
		cr.andEqualTo("mealApplicationId",id);
		List<MealTransaction> transactionList = transactionMapper.selectByExample(ex);
		SimpleDateFormat sdf=new SimpleDateFormat("yyyy/mm/dd HH:mm:ss");
		Map<String, Object> map = new HashMap<String, Object>();
		map.put("name", application.getName());
		map.put("unit",application.getUnit() );
		map.put("phone", application.getPhone());
		map.put("purpose", application.getPurpose());
		map.put("number", application.getNumber());
		map.put("createTime", sdf.format(application.getCreateTime()));
		if(approvalList.size()>0) {
			map.put("approver", approvalList.get(0).getApprover());
			if(0 == approvalList.get(0).getStatus()) {
				map.put("status", "驳回");
			}else if(1 == approvalList.get(0).getStatus()) {
				map.put("status", "通过");
			}
			map.put("reason",approvalList.get(0).getReason() );
			map.put("approvalTime",sdf.format(approvalList.get(0).getApprovalTime()));
		}
		if(transactionList.size()>0) {
			map.put("cardTransactor", transactionList.get(0).getTransactor());
			map.put("cardTransactionTime",sdf.format(transactionList.get(0).getTransactionTime()));
			map.put("cardResult","完成");
		
		}
			WordUtil.export(request, response, map, "", String.format("%s-chuchai.doc", application.getName()));
	}
	
	
}
