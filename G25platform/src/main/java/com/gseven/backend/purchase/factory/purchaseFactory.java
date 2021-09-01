package com.gseven.backend.purchase.factory;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.dao.EmptyResultDataAccessException;
import org.springframework.jdbc.core.BeanPropertyRowMapper;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.RowMapper;

import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.gseven.backend.hr.factory.HRFactory;
import com.gseven.backend.sys.entity.Department;
import com.gseven.backend.sys.entity.SysUser;
import com.gseven.backend.sys.service.ExcelUtil;
import com.gseven.backend.sys.service.TextUtil;

public class purchaseFactory {

	private static final Logger log = LoggerFactory.getLogger(HRFactory.class);

	public purchaseFactory() {
	}

	@SuppressWarnings("finally")
	public int SelectIsValue1(HttpServletRequest request, JdbcTemplate jdbcTemplate) {

		int serialNumber = 0;
		List<Map<String, Object>> list = null;
		try {

			String sql = "Select IMA01 from G25DB.PURCHASE_PSI_VALUE where COM='666' and IMA01='"
					+ request.getParameter("ima01") + "'";

			list = jdbcTemplate.queryForList(sql);
			System.out.println("SQL: " + sql);

			if (list != null && !list.isEmpty()) {
				serialNumber = 1;
			}
		} catch (Exception ex) {
			System.out.println("資料庫連線例外：" + ex);
		} finally {
			return serialNumber;
		}
	}

	@SuppressWarnings("finally")
	public int insertValue1(HttpServletRequest request, JdbcTemplate jdbcTemplate) {

		int i = 0;
		try {
			String sql = "INSERT INTO G25DB.PURCHASE_PSI_VALUE (COM,IMA01,PS ) " + "VALUES ('666','"
					+ request.getParameter("ima01") + "','" + request.getParameter("ps") + "')";

			System.out.println("SQL: " + sql);
			i = jdbcTemplate.update(sql);

		} catch (Exception ex) {
			System.out.println("資料庫連線例外 : " + ex);

		} finally {
			return i;
		}

	}

	@SuppressWarnings("finally")
	public int updateValue1(HttpServletRequest request, JdbcTemplate jdbcTemplate) {

		int i = 0;

		try {
			String sql = "update G25DB.PURCHASE_PSI_VALUE set PS='" + request.getParameter("ps")
					+ "' where COM='666' and IMA01='" + request.getParameter("ima01") + "'";

			System.out.println("SQL: " + sql);

			i = jdbcTemplate.update(sql);

		} catch (Exception ex) {
			System.out.println("資料庫連線例外 : " + ex);

		} finally {
			return i;
		}

	}

	@SuppressWarnings("finally")
	public int SelectIsValue(HttpServletRequest request, int i, JdbcTemplate jdbcTemplate) {

		Calendar c = Calendar.getInstance();
		c.add(Calendar.MONTH, i - 1);
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM");
		String d1 = format.format(c.getTime());

		int serialNumber = 0;
		List<Map<String, Object>> list = null;
		try {

			String sql = "Select IMA01 from G25DB.PURCHASE_PSI where COM='666' and IMA01='"
					+ request.getParameter("ima01") + "' and MONTH='" + d1 + "'";

			list = jdbcTemplate.queryForList(sql);
			System.out.println("SQL: " + sql);

			if (list != null && !list.isEmpty()) {
				serialNumber = 1;
			}
		} catch (Exception ex) {
			System.out.println("資料庫連線例外：" + ex);
		} finally {
			return serialNumber;
		}

	}

	@SuppressWarnings("finally")
	public int insertValue(HttpServletRequest request, int i, JdbcTemplate jdbcTemplate) {

		Calendar c = Calendar.getInstance();
		c.add(Calendar.MONTH, i - 1);
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM");
		String d1 = format.format(c.getTime());

		try {
			String sql = "INSERT INTO G25DB.PURCHASE_PSI (COM,IMA01,MONTH,VALUE ) " + "VALUES ('666','"
					+ request.getParameter("ima01") + "','" + d1 + "','" + request.getParameter("set_d" + i) + "')";

			System.out.println("SQL: " + sql);
			i = jdbcTemplate.update(sql);

		} catch (Exception ex) {
			System.out.println("資料庫連線例外 : " + ex);

		} finally {
			return i;
		}
	}

	@SuppressWarnings("finally")
	public int updateValue(HttpServletRequest request, int i, JdbcTemplate jdbcTemplate) {

		Calendar c = Calendar.getInstance();
		c.add(Calendar.MONTH, i - 1);
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM");
		String d1 = format.format(c.getTime());

		try {
			String sql = "update G25DB.PURCHASE_PSI set value='" + request.getParameter("set_d" + i)
					+ "' where COM='666' and IMA01='" + request.getParameter("ima01") + "' and MONTH='" + d1 + "' ";

			System.out.println("SQL: " + sql);

			i = jdbcTemplate.update(sql);

		} catch (Exception ex) {
			System.out.println("資料庫連線例外 : " + ex);

		} finally {
			return i;
		}
	}

	// PSI
	@SuppressWarnings("finally")
	public List<Map<String, Object>> selectIMA(HttpServletRequest request, JdbcTemplate jdbcTemplate) {

		String find_tqa = request.getParameter("find_tqa");
		String find_IMA01 = request.getParameter("find_IMA01");
		String order_is = request.getParameter("order_is");
		String can_get_order = request.getParameter("can_get_order");
		String type = request.getParameter("type");

		String find_pmc = request.getParameter("find_pmc");
		String find_ima54 = request.getParameter("find_ima54");
		String find_IMA131 = request.getParameter("find_IMA131");
		String find_ima1004 = request.getParameter("find_ima1004");

		String find_IMA06 = request.getParameter("find_IMA06");

		List<Map<String, Object>> list = null;

		Calendar c = Calendar.getInstance();
		c.add(Calendar.MONTH, 0);
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM");
		String d1 = format.format(c.getTime());


		c = Calendar.getInstance();
		c.add(Calendar.MONTH, 1);
		format = new SimpleDateFormat("yyyy-MM");
		String d2 = format.format(c.getTime());

		format = new SimpleDateFormat("M月 -yy");
		String dd2 = format.format(c.getTime());

		c = Calendar.getInstance();
		c.add(Calendar.MONTH, 2);
		format = new SimpleDateFormat("yyyy-MM");
		String d3 = format.format(c.getTime());

		format = new SimpleDateFormat("M月 -yy");
		String dd3 = format.format(c.getTime());

		c = Calendar.getInstance();
		c.add(Calendar.MONTH, 3);
		format = new SimpleDateFormat("yyyy-MM");
		String d4 = format.format(c.getTime());

		format = new SimpleDateFormat("M月 -yy");
		String dd4 = format.format(c.getTime());

		c = Calendar.getInstance();
		c.add(Calendar.MONTH, 4);
		format = new SimpleDateFormat("yyyy-MM");
		String d5 = format.format(c.getTime());

		format = new SimpleDateFormat("M月 -yy");
		String dd5 = format.format(c.getTime());

		c = Calendar.getInstance();
		c.add(Calendar.MONTH, 5);
		format = new SimpleDateFormat("yyyy-MM");
		String d6 = format.format(c.getTime());

		format = new SimpleDateFormat("M月 -yy");
		String dd6 = format.format(c.getTime());
		try {

			String sql = " select \r\n" + " aa, \r\n" + " bb, \r\n" + " cc, \r\n"
					+ "NVL((select sum(value) from G25DB.PURCHASE_PSI where COM='666' and PURCHASE_PSI.IMA01=a.IMA01 and month ='"+d1+"'),0)  d1, \r\n"
					+ "NVL((select sum(value) from G25DB.PURCHASE_PSI where COM='666' and PURCHASE_PSI.IMA01=a.IMA01 and month ='"+d2+"'),0)  d2, \r\n"
					+ "NVL((select sum(value) from G25DB.PURCHASE_PSI where COM='666' and PURCHASE_PSI.IMA01=a.IMA01 and month ='"+d3+"'),0)  d3, \r\n"
					+ "NVL((select sum(value) from G25DB.PURCHASE_PSI where COM='666' and PURCHASE_PSI.IMA01=a.IMA01 and month ='"+d4+"'),0)  d4, \r\n"
					+ "NVL((select sum(value) from G25DB.PURCHASE_PSI where COM='666' and PURCHASE_PSI.IMA01=a.IMA01 and month ='"+d5+"'),0)  d5, \r\n"
					+ "NVL((select sum(value) from G25DB.PURCHASE_PSI where COM='666' and PURCHASE_PSI.IMA01=a.IMA01 and month ='"+d6+"'),0)  d6, \r\n"
					+ "NVL((select sum(OEB12-OEB24+OEB25-OEB26) from DB_G.OEB_FILE where  oeb04=a.IMA01 and OEB70 = 'N'  and (OEB12-OEB24+OEB25-OEB26) > 0 and  oeb15 like '%"+dd2+"%'),0) o1 , \r\n"
					+ "NVL((select sum(OEB12-OEB24+OEB25-OEB26) from DB_G.OEB_FILE where  oeb04=a.IMA01 and OEB70 = 'N'  and (OEB12-OEB24+OEB25-OEB26) > 0 and  oeb15 like '%"+dd3+"%'),0) o2 , \r\n"
					+ "NVL((select sum(OEB12-OEB24+OEB25-OEB26) from DB_G.OEB_FILE where  oeb04=a.IMA01 and OEB70 = 'N'  and (OEB12-OEB24+OEB25-OEB26) > 0 and   oeb15 like '%"+dd4+"%'),0) o3 , \r\n"
					+ "NVL((select sum(OEB12-OEB24+OEB25-OEB26) from DB_G.OEB_FILE where  oeb04=a.IMA01 and OEB70 = 'N'  and (OEB12-OEB24+OEB25-OEB26) > 0 and   oeb15 like '%"+dd5+"%'),0) o4 , \r\n"
					+ "NVL((select sum(OEB12-OEB24+OEB25-OEB26) from DB_G.OEB_FILE where  oeb04=a.IMA01 and OEB70 = 'N'  and (OEB12-OEB24+OEB25-OEB26) > 0 and  oeb15 like '%"+dd6+"%'),0) o5 , \r\n"
					+ "NVL((select sum(OEB12-OEB24+OEB25-OEB26) from DB_G.OEB_FILE where  oeb04=a.IMA01 and OEB70 = 'N'  and (OEB12-OEB24+OEB25-OEB26) > 0 and OEB15 < LAST_DAY(SYSDATE)  ),0) a1 , \r\n"//2.6	售訂未銷量
					//+ "NVL((select sum(oeb12) from DB_G.OEB_FILE where  oeb04=a.IMA01 and OEB70 = 'N'  and OEB15 < LAST_DAY(SYSDATE) ),0) a3 , \r\n"
					//+ "NVL((select sum(OEB23) from DB_G.OEB_FILE where  oeb04=a.IMA01   and OEB70 = 'N' and OEB15 < LAST_DAY(SYSDATE) ),0) a5 , \r\n"
					//+ "NVL((select sum(oeb24) from DB_G.OEB_FILE where  oeb04=a.IMA01   and OEB70 = 'N' and   OEB15 < LAST_DAY(SYSDATE)  ),0) a6 , \r\n"
					//+ "NVL((select sum(oeb25) from DB_G.OEB_FILE where  oeb04=a.IMA01   and OEB70 = 'N' and    OEB15 < LAST_DAY(SYSDATE)  ),0) a7 , \r\n"
					//+ "NVL((select sum(oeb26) from DB_G.OEB_FILE where  oeb04=a.IMA01   and OEB70 = 'N' and    OEB15 < LAST_DAY(SYSDATE)  ),0) a8 , \r\n"
					+ "(select max(ps) from G25DB.PURCHASE_PSI_VALUE where  COM='666' and PURCHASE_PSI_VALUE.IMA01=a.IMA01)  ps,     \r\n"
					+ "( select SUM(OEB12-OEB24+OEB25-OEB26) from db_g.OEB_FILE where oeb04=a.ima01 and OEB70 = 'N' and (OEB12-OEB24+OEB25-OEB26) > 0 )  SUMOEB12,-- 品號總受訂量,  \r\n"
					+ "        NVL((SELECT DISTINCT SUM(IMG10) OVER (PARTITION BY IMG01,IMG03) FROM DB_G.IMG_FILE G1 WHERE IMG02 LIKE '%S' AND IMG03 = '備貨' AND G1.IMG01 = a.ima01),0) \r\n"
					+ "        SUMIMG10IMG03_1, --存貨倉備貨庫存總量(備貨數量a1) \r\n"
					+ "        NVL((SELECT DISTINCT SUM(IMG10) OVER (PARTITION BY IMG01,IMG03) FROM DB_G.IMG_FILE G2 WHERE IMG02 LIKE '%S' AND IMG03 = '陳列' AND G2.IMG01 = a.ima01),0) \r\n"
					+ "        SUMIMG10IMG03_2, --存貨倉陳列庫存總量(陳列數量a2) \r\n"
					+ "        NVL((SELECT DISTINCT SUM(IMG10) OVER (PARTITION BY IMG01) FROM DB_G.IMG_FILE G5 WHERE IMG02 LIKE '%S' AND G5.IMG01 = a.ima01),0) \r\n"
					+ "        SUMIMG10IMG02_1, --存貨倉庫存總量 \r\n"
					+ "        NVL((SELECT DISTINCT SUM(IMG10) OVER (PARTITION BY IMG01,IMG03) FROM DB_G.IMG_FILE G3 WHERE IMG02 LIKE '%C' AND IMG03 = '備貨' AND G3.IMG01 = a.ima01),0) \r\n"
					+ "        SUMIMG10IMG03_3, --承銷倉備貨庫存總量 \r\n"
					+ "        NVL((SELECT DISTINCT SUM(IMG10) OVER (PARTITION BY IMG01,IMG03) FROM DB_G.IMG_FILE G4 WHERE IMG02 LIKE '%C' AND IMG03 = '陳列' AND G4.IMG01 =a.ima01),0) \r\n"
					+ "        SUMIMG10IMG03_4, --承銷倉陳列庫存總量 \r\n"
					+ "        NVL((SELECT DISTINCT SUM(IMG10) OVER (PARTITION BY IMG01) FROM DB_G.IMG_FILE G6 WHERE IMG02 LIKE '%C' AND G6.IMG01 = a.ima01),0) \r\n"
					+ "        SUMIMG10IMG02_2, --承銷倉庫存總量(承銷數量a3) \r\n"
					+ "        NVL((SELECT DISTINCT SUM(IMG10) OVER (PARTITION BY IMG01) FROM DB_G.IMG_FILE G7 WHERE IMG02 LIKE '%W' AND G7.IMG01 = a.ima01),0) \r\n"
					+ "        SUMIMG10IMG02_3, --在途倉庫存總量 \r\n"
					+ "        NVL((SELECT DISTINCT SUM(IMG10) OVER (PARTITION BY IMG01) FROM DB_G.IMG_FILE G8 WHERE IMG02 NOT LIKE '%W' AND G8.IMG03 <> '備貨' AND G8.IMG03 <> '陳列' AND G8.IMG01 =a.ima01),0) \r\n"
					+ "        SUMIMG10IMG02_4, --其他倉庫儲位庫存總量 \r\n"
					+ "( select SUM(IMG10)  from DB_G.img_file where img01=a.ima01) SUMIMG10 --品號總量-- \r\n"
					+ ",a.* \r\n" + ",t1.tqa02 tqa02t1" + ",t2.tqa02 tqa02t2" + ",imz_file.* \r\n" + ",oba_file.* \r\n" + ",PMC_file.* \r\n"
					+ "from db_g.ima_FILE  a --訂單單身\r\n" 
					+ "left join db_g.tqa_file t1 on t1.tqa01 = a.ima1005   --品牌\r\n"
					+ "left join db_g.imz_file on imz01 = a.ima06 --採購大類\r\n"
					+ "left join db_g.oba_file on oba01 = a.ima131 --商品大類\r\n"
					+ "LEFT JOIN DB_G.PMC_FILE ON PMC01 = a.IMA54     --供應商\r\n"
					+ "LEFT JOIN DB_G.tqa_FILE t2 ON t2.tqa01 = a.IMA1004     --商品狀態\r\n"
					+ "left join ( select OEB04,SUM(OEB12-OEB24+OEB25-OEB26) aa from db_g.OEB_FILE where  OEB70 = 'N' and (OEB12-OEB24+OEB25-OEB26) > 0 group by oeb04 ) k on k.oeb04=a.ima01 \r\n"
					+ "left join (SELECT DISTINCT SUM(IMG10) OVER (PARTITION BY IMG01,IMG03) bb,img01 FROM DB_G.IMG_FILE WHERE IMG02 LIKE '%S' AND IMG03 = '備貨' ) m on m.img01=a.ima01 \r\n"
					+ "left join (select OEB04,sum(OEB12-OEB24+OEB25-OEB26) cc from DB_G.OEB_FILE where  OEB70 = 'N'  and (OEB12-OEB24+OEB25-OEB26) > 0 and OEB15 < LAST_DAY(SYSDATE) group by oeb04) n on n.OEB04=a.ima01 \r\n";

			List<Object> paramList = new ArrayList<>();
			Object[] params = null;
			String where = " WHERE 1 = 1  and imaacti='Y' and  ima01 not like 'MISC%' and  ima01 not like 'GZ%' and ima01 not like 'V%'  \r\n";

			if ("list".equals(type)) {
				where += " and rownum <=300 ";
			}
			// 品牌
			if (!TextUtil.isBlankOrNull(find_tqa)) {
				where += " AND a.IMA1005 =?  \r\n";
				paramList.add(find_tqa);
			}
			// 品號
			if (!TextUtil.isBlankOrNull(find_IMA01)) {
				where += " AND a.IMA01 =?  \r\n";
				paramList.add(find_IMA01);
			}

			// 供應商
			if (!TextUtil.isBlankOrNull(find_pmc)) {
				where += " AND a.IMA54 =?  \r\n";
				paramList.add(find_pmc);
			}

			// 產品狀態
			if (!TextUtil.isBlankOrNull(find_ima1004)) {
				where += " AND a.IMA1004 =?  \r\n";
				paramList.add(find_ima1004);
			}

			// 採購大類
			if (!TextUtil.isBlankOrNull(find_IMA06)) {
				where += " AND a.IMA06 =?  \r\n";
				paramList.add(find_IMA06);
			}

			// 產品大類
			if (!TextUtil.isBlankOrNull(find_IMA131)) {
				where += " AND a.IMA131 =?  \r\n";
				paramList.add(find_IMA131);
			}

			// NVL((select sum(OEB12-OEB24+OEB25-OEB26) from DB_G.OEB_FILE where
			// oeb04=a.IMA01 and OEB70 = 'N' and (OEB12-OEB24+OEB25-OEB26) > 0 and OEB15 <
			// LAST_DAY(SYSDATE) ),0)
			if (!TextUtil.isBlankOrNull(order_is)) {
				where += "  and aa>0 ";
			}
			if (!TextUtil.isBlankOrNull(can_get_order)) {
				if ("2".equals(can_get_order)) {
					where += " AND NVL(bb,0)-NVL(cc,0)>0  \r\n";
				} else {
					where += " AND NVL(bb,0)-NVL(cc,0) <=0  \r\n";
				}

			}

			sql += where;
			params = paramList.toArray(new Object[paramList.size()]);

			System.out.println("SQL: " + TextUtil.formatSqlStatmentDateFull(sql, params));

			list = jdbcTemplate.queryForList(sql, params);

		} catch (Exception ex) {
			System.out.println("資料庫連線例外：" + ex);
			log.info("資料庫連線例外：" + ex);
		} finally {
			return list;
		}

	}

	@SuppressWarnings("finally")
	public List<Map<String, Object>> selectTQA(JdbcTemplate jdbcTemplate) {
		List<Map<String, Object>> list = new ArrayList<>();
		try {

			String sql = "select * from DB_G.TQA_FILE where tqa03='2' order by tqa01";

			// System.out.println(sql);
			list = jdbcTemplate.queryForList(sql);

		} catch (Exception ex) {
			ex.printStackTrace();
			System.out.println("資料庫連線例外：" + ex);
			log.info("資料庫連線例外：" + ex);
		} finally {
			return list;
		}
	}

	@SuppressWarnings("finally")
	public List<Map<String, Object>> selectDepartment(String loginId, JdbcTemplate jdbcTemplate) {
		List<Map<String, Object>> list = new ArrayList<>();
		try {
			String sql = "SELECT dep.CODE  " + "FROM G25DB.employee emp\r\n"
					+ "INNER JOIN G25DB.department dep ON emp.departmentid=dep.departmentid \r\n"
					+ "WHERE emp.code=? AND emp.flag='T'";

			List<Object> paramList = new ArrayList<>();

			Object[] params = null;
			paramList.add(loginId);

			params = paramList.toArray(new Object[paramList.size()]);
			list = jdbcTemplate.queryForList(sql, params);

			// System.out.println(sql + ", "+ departmentId);
		} catch (Exception ex) {
			ex.printStackTrace();
			System.out.println("資料庫連線例外：" + ex);
			log.info("資料庫連線例外：" + ex);
		} finally {
			return list;
		}

	}

}