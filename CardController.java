package com.cplatform.operator.auth.controller;


import java.text.SimpleDateFormat;
import java.util.ArrayList;
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
import com.cplatform.app.domain.CardCancelApproval;
import com.cplatform.app.domain.CardCancelTransaction;
import com.cplatform.app.domain.CardTransaction;
import com.cplatform.app.domain.TuserCard;
import com.cplatform.app.mapper.CardApplicationMapper;
import com.cplatform.app.mapper.CardApprovalMapper;
import com.cplatform.app.mapper.CardCancelApplicationMapper;
import com.cplatform.app.mapper.CardCancelApprovalMapper;
import com.cplatform.app.mapper.CardCancelTransactionMapper;
import com.cplatform.app.mapper.CardTransactionMapper;
import com.cplatform.app.mapper.TuserCardMapper;
import com.cplatform.app.service.CardApplicationService;
import com.cplatform.app.service.CardCancelApplicationService;
import com.cplatform.app.util.ExportExcel;
import com.cplatform.app.util.ImgUtils;
import com.github.pagehelper.PageHelper;
import com.github.pagehelper.PageInfo;

import tk.mybatis.mapper.common.Mapper;
import tk.mybatis.mapper.entity.Example;
import tk.mybatis.mapper.entity.Example.Criteria;
/**
 * 卡卷办理和注销
 * @author ZhouLe
 *
 */
@Controller
@RequestMapping("card")
public class CardController extends BaseController<CardApplication, CardApplication> {
	
	@Autowired
	private CardApplicationMapper mapper;
	@Autowired
	private CardApprovalMapper approvalMapper;
	@Autowired
	private CardTransactionMapper transactionMapper;
	@Autowired
	private CardApplicationService service;
	@Autowired
	private CardCancelApplicationService cancelService;
	@Autowired
	private CardCancelApplicationMapper cancelMapper;
	@Autowired
	private CardCancelApprovalMapper cancelApprovalMapper;
	@Autowired
	private CardCancelTransactionMapper cancelTransactionMapper;
	@Autowired
	private TuserCardMapper tuserCardMapper;
	@Override
	public Mapper<CardApplication> getMapper() {
		return mapper;
	}
	/**
     * 卡卷申请列表
     * @param approval
     * @param page
     * @return
     */ 
	@RequestMapping("list")
	public ModelAndView list(CardApplication application, Page page) {
		ModelAndView mv = new ModelAndView("card/cardList");
		PageHelper.startPage(page.getPageNo(), page.getPageSize());
		List<CardApplication> list = service.FindAll(application);
		PageInfo<CardApplication> pageList = new PageInfo<>(list);
    	mv.addObject("page",pageList);
        mv.addObject("cur_nav", "cardList");
    	return mv;
	}
	/**
     * 卡卷注销申请列表
     * @param approval
     * @param page
     * @return
     */
	@RequestMapping("cardCancelList")
	public ModelAndView cancelList(CardCancelApplication application, Page page) {
		ModelAndView mv = new ModelAndView("card/cardCancelList");
		PageHelper.startPage(page.getPageNo(), page.getPageSize());
		List<CardCancelApplication> list = cancelService.FindAll(application);
		PageInfo<CardCancelApplication> pageList = new PageInfo<>(list);
    	mv.addObject("page",pageList);
        mv.addObject("cur_nav", "cancelList");
    	return mv;
	}
	/**
     * 餐卡；门禁卡申请详细数据
     * @param approval
     * @param page
     * @return
     */
	@RequestMapping("selectOne")
	public ModelAndView SelectOne(Long id) {
		ModelAndView mv = new ModelAndView("card/listDetail");
		//查询申请表信息
		CardApplication cardApplication = mapper.selectByPrimaryKey(id);
		List<CardApplication> list = new ArrayList<CardApplication>();
		List<CardApproval> approvalList = new ArrayList<>();
		List<CardTransaction> transactionList = new ArrayList<>();
		List<CardTransaction> transactionListDoor = new ArrayList<>();
			//查询餐卡申请信息
			Example exa = new Example(CardApplication.class);
			Criteria cri = exa.createCriteria();
			cri.andEqualTo("card_id",cardApplication.getCardId());
			cri.andEqualTo("type",0);
			if (cardApplication.getStatus() == 0) {
				cri.andEqualTo("status",0);
			}
			list = mapper.selectByExample(exa);//餐卡信息办理
			//查询餐卡审批表信息
			Example example = new Example(CardApproval.class);
			Criteria criteria = example.createCriteria();
			criteria.andEqualTo("applicationId",list.get(0).getId());
			approvalList = approvalMapper.selectByExample(example);
			//查询餐卡办理表信息
			Example ex = new Example(CardTransaction.class);
			Criteria cr = example.createCriteria();
			cr.andEqualTo("applicationId",list.get(0).getId());
			 transactionList = transactionMapper.selectByExample(ex);
			 //查询门禁卡办理信息
			 Example exDoor = new Example(CardTransaction.class);
			 Criteria crDoor = exDoor.createCriteria();
			 crDoor.andEqualTo("applicationId",id);
			 transactionListDoor = transactionMapper.selectByExample(ex);
		String[] photoUrl = cardApplication.getProvePhotoUrl().split(";");
		
		
		if(approvalList.size() > 0) {
			mv.addObject("approvalList",approvalList.get(0));
		}
		if(transactionList.size() > 0) {
			mv.addObject("transactionList",transactionList.get(0));
		}
		if(transactionListDoor.size() > 0) {
			mv.addObject("transactionListDoor",transactionListDoor.get(0));//门禁卡办理信息
		}
    	mv.addObject("cardApproval",cardApplication);//申请表信息
    	mv.addObject("photoUrl",photoUrl);//其他材料
        mv.addObject("cur_nav", "selectOne");
    	return mv;
	}
	
	/**
     * 通行证申请详细数据
     * @param approval
     * @param page
     * @return
     */
	@RequestMapping("selectCar")
	public ModelAndView SelectCar(Long id) {
		ModelAndView mv = new ModelAndView("card/listDetailCar");
		//查询申请表信息
		CardApplication cardApplication = mapper.selectByPrimaryKey(id);
		List<CardApproval> approvalList = new ArrayList<>();
		List<CardTransaction> transactionList = new ArrayList<>();
			//查询审批表信息
			Example example = new Example(CardApproval.class);
			Criteria criteria = example.createCriteria();
			criteria.andEqualTo("applicationId",id);
			 approvalList = approvalMapper.selectByExample(example);
			//查询办理表信息
			Example ex = new Example(CardTransaction.class);
			Criteria cr = example.createCriteria();
			cr.andEqualTo("applicationId",id);
			 transactionList = transactionMapper.selectByExample(ex);
		String[] photoUrl = cardApplication.getProvePhotoUrl().split(";");
		
		
		if(approvalList.size() > 0) {
			mv.addObject("approvalList",approvalList.get(0));
		}
		if(transactionList.size() > 0) {
			mv.addObject("transactionList",transactionList.get(0));
		}
		
    	mv.addObject("cardApproval",cardApplication);//申请表信息
    	mv.addObject("photoUrl",photoUrl);//其他材料
        mv.addObject("cur_nav", "selectOne");
    	return mv;
	}
	
	/**
	 * 卡卷注销申请详细数据
	 * @param approval
	 * @param page
	 * @return
	 */
	@RequestMapping("selectCancelOne")
	public ModelAndView SelectCancelOne(Long id) {
		
		List<CardCancelApproval> approvalList = new ArrayList<>();
		List<CardCancelTransaction> transactionList = new ArrayList<>();
		//查询申请表信息
		CardCancelApplication cardApplication = cancelMapper.selectByPrimaryKey(id);
		//查询审批表信息
		Example example = new Example(CardCancelApproval.class);
		Criteria criteria = example.createCriteria();
		criteria.andEqualTo("cancelApplicationId",id);
		//查询办理表信息
		Example ex = new Example(CardTransaction.class);
		Criteria cr = example.createCriteria();
		cr.andEqualTo("cancelApplicationId",id);
		//判断是否为餐卡申请
		CardCancelApproval mealAudit = new CardCancelApproval();//将审核信息放入新建对象中
		if(null != cardApplication.getType() && cardApplication.getType() == 0) {
			ModelAndView mv = new ModelAndView("card/cancelMealDetail");//跳转页面
			//若是餐卡申请则查询餐卡审核
			approvalList = cancelApprovalMapper.selectByExample(example);
			for (CardCancelApproval cancelApproval : approvalList) {
				//判断是餐卡审核还是餐卡审批
				if(null != cancelApproval.getMealApprovalType() && cancelApproval.getMealApprovalType() == 0) {
					 mealAudit = cancelApproval;
					 approvalList.remove(cancelApproval);//集合中删除该纪录
				}
			}
			transactionList = cancelTransactionMapper.selectByExample(ex);
			for (CardCancelTransaction transaction : transactionList) {
				//判断是餐卡办理还是财务办理
				if(null != transaction.getType() && transaction.getType() == 2 ) {
					transactionList.remove(transaction);
				}
			}
			
			
			if(approvalList.size() > 0) {
				mv.addObject("approvalList",approvalList.get(0));
			}
			if(transactionList.size() > 0) {
				mv.addObject("transactionList",transactionList.get(0));
			}
			mv.addObject("cardApproval",cardApplication);
			mv.addObject("mealAudit",mealAudit);//餐卡审核
			mv.addObject("cur_nav", "cancelList");
			return mv;
		}else {
			ModelAndView mv = new ModelAndView("card/cancelDetail");//跳转页面
			//审批表
			approvalList = cancelApprovalMapper.selectByExample(example);
			//办理表
			transactionList = cancelTransactionMapper.selectByExample(ex);
			
			
			if(approvalList.size() > 0) {
				mv.addObject("approvalList",approvalList.get(0));
			}
			if(transactionList.size() > 0) {
				mv.addObject("transactionList",transactionList.get(0));
			}
			mv.addObject("cardApproval",cardApplication);
			mv.addObject("cur_nav", "cancelList");
			return mv;
		}
		
		
	}
	/**
	 * 卡卷申请信息导出excel
	 */
	@RequestMapping("exportWorkExcel")
	public void exportExcel(CardApplication application,HttpServletRequest request,HttpServletResponse response)throws Exception {
		List<CardApplication> list = service.FindAll(application);
		//2.导出
		//这里设置的文件格式是application/x-excel
		response.setContentType("application/x-excel");
		String fileName ="卡卷申请信息.xls";
		response.setHeader("Content-Disposition", "attachment;filename=" + new String(fileName.getBytes(), "ISO-8859-1"));
		ServletOutputStream outputStream = response.getOutputStream();
		ExportExcel.exportWorkExcel(list, outputStream);
		if(outputStream != null) {
            outputStream.close();
        }
	}
	/**
	 * 卡卷注销申请信息导出excel
	 */
	@RequestMapping("exportCancelExcel")
	public void exportCancelExcel(CardCancelApplication application,HttpServletRequest request,HttpServletResponse response)throws Exception {
		List<CardCancelApplication> list = cancelService.FindAll(application);
		//2.导出
		//这里设置的文件格式是application/x-excel
		response.setContentType("application/x-excel");
		String fileName ="卡卷注销申请信息.xls";
		response.setHeader("Content-Disposition", "attachment;filename=" + new String(fileName.getBytes(), "ISO-8859-1"));
		ServletOutputStream outputStream = response.getOutputStream();
		
		ExportExcel.exportCancelExcel(list, outputStream);
		if(outputStream != null) {
			outputStream.close();
		}
	}
	/***
	 * 导出餐卡、门禁卡申请详细数据
	 * 
	 * @return
	 * @throws Exception
	 */
	@RequestMapping(value = "cardExport", method = RequestMethod.GET)
	@ResponseBody
	public void cardExport(HttpServletRequest request, HttpServletResponse response, Long id) throws Exception {
		String realPath = request.getSession().getServletContext().getRealPath("/");
		String path = realPath + "/tmp/zip/";
		String excelPath = realPath + "/tmp/excel/";
		String dirOrFile = realPath + "/tmp/excel/";
		//查询申请表信息
		CardApplication cardApplication = mapper.selectByPrimaryKey(id);
		//查询用户表信息
		Example example1 = new Example(TuserCard.class);
		Criteria criteria1 = example1.createCriteria();
		criteria1.andEqualTo("id",cardApplication.getUserId());
		List<TuserCard> tUserCardList = tuserCardMapper.selectByExample(example1);
		
		List<CardApplication> list = new ArrayList<CardApplication>();
		List<CardApproval> approvalList = new ArrayList<>();
		List<CardTransaction> transactionList = new ArrayList<>();
		List<CardTransaction> transactionListDoor = new ArrayList<>();
		//判断是否是门禁卡办理;如果是则查询餐卡办理
		if(cardApplication.getType() == 1) {
			Example exa = new Example(CardApplication.class);
			Criteria cri = exa.createCriteria();
			cri.andEqualTo("card_id",cardApplication.getCardId());
			cri.andEqualTo("type",0);
			if (cardApplication.getStatus() == 0) {
				cri.andEqualTo("status",0);
			}
			list = mapper.selectByExample(exa);//餐卡申请信息
			//查询餐卡审批表信息
			Example example = new Example(CardApproval.class);
			Criteria criteria = example.createCriteria();
			criteria.andEqualTo("applicationId",list.get(0).getId());
			approvalList = approvalMapper.selectByExample(example);
			//查询餐卡办理表信息
			Example ex = new Example(CardTransaction.class);
			Criteria cr = example.createCriteria();
			cr.andEqualTo("applicationId",list.get(0).getId());
			 transactionList = transactionMapper.selectByExample(ex);
			 //查询门禁卡办理信息
			 Example exDoor = new Example(CardTransaction.class);
			 Criteria crDoor = exDoor.createCriteria();
			 crDoor.andEqualTo("applicationId",id);
			 transactionListDoor = transactionMapper.selectByExample(ex);
		}
		
		SimpleDateFormat sdf=new SimpleDateFormat("yyyy/mm/dd HH:mm:ss");
			Map<String, Object> map = new HashMap<String, Object>();
			map.put("name", cardApplication.getName());
			map.put("unit",cardApplication.getUnit() );
			map.put("cardId",cardApplication.getCardId() );
			map.put("phone", cardApplication.getPhone());
			map.put("createTime", sdf.format(cardApplication.getCreateTime()));
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
				//餐卡信息
					map.put("cardTransactor", transactionList.get(0).getTransactor());
					map.put("cardTransactionTime",sdf.format(transactionList.get(0).getTransactionTime()));
					map.put("mealCard",tUserCardList.get(0).getMealCard());
					map.put("cardResult","完成");
				}
			if(transactionListDoor.size()>0) {
				//门禁卡
				map.put("mealTransactor", transactionListDoor.get(0).getTransactor());
				map.put("transactionTime",sdf.format(transactionListDoor.get(0).getTransactionTime()));
				map.put("cardKey",tUserCardList.get(0).getCardKey());
				map.put("mealResult","完成");
			}
		  WordUtil.export(request, response, map, "cardAndMeal.ftl", String.format("%s-餐卡、门禁卡申请.doc", cardApplication.getName()));
		//导出图片
		  //一寸照片
		ImgUtils.saveImg(cardApplication.getIssunPhotoUrl());
		//其他证明材料
		String[] provePhotoUrls = cardApplication.getProvePhotoUrl().split(";");
		for (String provePhotoUrl : provePhotoUrls) {
			ImgUtils.saveImg(provePhotoUrl);
		}
	}
	/***
	 * 导出餐卡注销申请详细数据
	 * 
	 * @return
	 * @throws Exception
	 */
	@RequestMapping(value = "cardCancleExport", method = RequestMethod.GET)
	@ResponseBody
	public void cardCancleExport(HttpServletRequest request, HttpServletResponse response, Long id) throws Exception {
		String realPath = request.getSession().getServletContext().getRealPath("/");
		String path = realPath + "/tmp/zip/";
		String excelPath = realPath + "/tmp/excel/";
		String dirOrFile = realPath + "/tmp/excel/";
		SimpleDateFormat sdf=new SimpleDateFormat("yyyy/mm/dd HH:mm:ss");
		Map<String, Object> map = new HashMap<String, Object>();
		//查询申请表信息
		CardCancelApplication cancelApplication = cancelMapper.selectByPrimaryKey(id);
		//查询审批表信息
		Example example = new Example(CardCancelApproval.class);
		Criteria criteria = example.createCriteria();
		criteria.andEqualTo("cancelApplicationId",id);
		List<CardCancelApproval> approvalList = cancelApprovalMapper.selectByExample(example);
		for (CardCancelApproval cancelApproval : approvalList) {
			//判断是餐卡审核还是餐卡审批
			if(null != cancelApproval.getMealApprovalType() && cancelApproval.getMealApprovalType() == 0) {
				//审核信息录入
				map.put("transactor", cancelApproval.getApprover());
				map.put("money",cancelApplication.getMealCardMoney());
				map.put("moneyTransactionTime",sdf.format(cancelApproval.getApprovalTime()));
				 approvalList.remove(cancelApproval);//集合中删除该纪录
			}
		}
		//查询办理表信息
		Example ex = new Example(CardCancelTransaction.class);
		Criteria cr = example.createCriteria();
		cr.andEqualTo("cancelApplicationId",id);
		List<CardCancelTransaction> transactionList = cancelTransactionMapper.selectByExample(ex);
		
			map.put("name", cancelApplication.getName());
			map.put("unit",cancelApplication.getUnit() );
			map.put("cardId",cancelApplication.getCardId() );
			map.put("phone", cancelApplication.getPhone());
			map.put("createTime", sdf.format(cancelApplication.getCreateTime()));
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
				for (CardCancelTransaction cancelTransaction : transactionList) {
					if(0 == cancelTransaction.getType()) {
						map.put("cancelTransactor", cancelTransaction.getTransactor());
						map.put("cancelResult", "完成");
						map.put("cancelTransactionTime", cancelTransaction.getTransactionTime());
						
					}//财务办理
					if(2 == cancelTransaction.getType()) {
						map.put("cancelMealTransactor", cancelTransaction.getTransactor());
						map.put("cancelMoney", "完成");
						map.put("cancelMoneyTransactionTime",cancelTransaction.getTransactionTime());
					}
				}
				
			}
		  WordUtil.export(request, response, map, "mealCancle.ftl", String.format("%s-餐卡注销.doc", cancelApplication.getName()));
		//导出图片
		//证明材料
		ImgUtils.saveImg(cancelApplication.getProvePhotoUrl());
	}
	/***
	 * 导出车辆通行证申请详细数据
	 * 
	 * @return
	 * @throws Exception
	 */
	@RequestMapping(value = "carExport", method = RequestMethod.GET)
	@ResponseBody
	public void carExport(HttpServletRequest request, HttpServletResponse response, Long id) throws Exception {
		String realPath = request.getSession().getServletContext().getRealPath("/");
		String path = realPath + "/tmp/zip/";
		String excelPath = realPath + "/tmp/excel/";
		String dirOrFile = realPath + "/tmp/excel/";
		//查询申请表信息
		CardApplication cardApplication = mapper.selectByPrimaryKey(id);
		//查询审批表信息
		Example example = new Example(CardApproval.class);
		Criteria criteria = example.createCriteria();
		criteria.andEqualTo("applicationId",id);
		List<CardApproval> approvalList = approvalMapper.selectByExample(example);
		//查询办理表信息
		Example ex = new Example(CardTransaction.class);
		Criteria cr = example.createCriteria();
		cr.andEqualTo("applicationId",id);
		List<CardTransaction> transactionList = transactionMapper.selectByExample(ex);
		SimpleDateFormat sdf=new SimpleDateFormat("yyyy/mm/dd HH:mm:ss");
			Map<String, Object> map = new HashMap<String, Object>();
			map.put("name", cardApplication.getName());
			map.put("unit",cardApplication.getUnit() );
			map.put("cardId",cardApplication.getCardId() );
			map.put("phone", cardApplication.getPhone());
			map.put("userCarNumber", cardApplication.getCarNumber());
			map.put("createTime", sdf.format(cardApplication.getCreateTime()));
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
				map.put("carNumber",transactionList.get(0).getCarNumber());
				map.put("cardResult","完成");
			
			}
		  WordUtil.export(request, response, map, "car.ftl", String.format("%s-通行证申请.doc", cardApplication.getName()));
		//导出图片
		  //一寸照片
		ImgUtils.saveImg(cardApplication.getIssunPhotoUrl());
		//其他证明材料
		String[] provePhotoUrls = cardApplication.getProvePhotoUrl().split(";");
		for (String provePhotoUrl : provePhotoUrls) {
			ImgUtils.saveImg(provePhotoUrl);
		}
	}
	/***
	 * 导出门禁卡，通行证注销申请详细数据
	 * 
	 * @return
	 * @throws Exception
	 */
	@RequestMapping(value = "carCancleExport", method = RequestMethod.GET)
	@ResponseBody
	public void carCancleExport(HttpServletRequest request, HttpServletResponse response, Long id) throws Exception {
		String realPath = request.getSession().getServletContext().getRealPath("/");
		String path = realPath + "/tmp/zip/";
		String excelPath = realPath + "/tmp/excel/";
		String dirOrFile = realPath + "/tmp/excel/";
		//查询申请表信息
		CardCancelApplication cancelApplication = cancelMapper.selectByPrimaryKey(id);
		//查询审批表信息
		Example example = new Example(CardCancelApproval.class);
		Criteria criteria = example.createCriteria();
		criteria.andEqualTo("cancelApplicationId",id);
		List<CardCancelApproval> approvalList = cancelApprovalMapper.selectByExample(example);
		//查询办理表信息
		Example ex = new Example(CardCancelTransaction.class);
		Criteria cr = example.createCriteria();
		cr.andEqualTo("cancelApplicationId",id);
		List<CardCancelTransaction> transactionList = cancelTransactionMapper.selectByExample(ex);
		SimpleDateFormat sdf=new SimpleDateFormat("yyyy/mm/dd HH:mm:ss");
			Map<String, Object> map = new HashMap<String, Object>();
			map.put("name", cancelApplication.getName());
			map.put("unit",cancelApplication.getUnit() );
			map.put("cardId",cancelApplication.getCardId() );
			map.put("phone", cancelApplication.getPhone());
			map.put("userCarNumber", cancelApplication.getCarNumber());
			map.put("createTime", sdf.format(cancelApplication.getCreateTime()));
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
		  WordUtil.export(request, response, map, "carCanale.ftl", String.format("%s-通行证注销.doc", cancelApplication.getName()));
		//导出图片
		  //证明材料
		ImgUtils.saveImg(cancelApplication.getProvePhotoUrl());
	}
}
