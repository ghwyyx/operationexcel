package com.gh.operateexcel;

public class ExcelDataVO {
	/**
     * ʡ��
     */
    private String province;

    /**
     * ����
     */
    private String consNo;

    /**
     * ��������
     */
    private String consSortCode;

    /**
     * Ƿ�Ѳ�ѯ
     */
    private String search;
   // �µ�
    private String saleOrder;
    // GOODS_CODE
    private String goodsCode;
    // ֧�����
    private String pay;
    // ֧��������
    private String orderNo;
    // ���˽��
    private String yxBackResult;

    public String getProvince() {
        return province;
    }

    public void setProvince(String province) {
        this.province = province;
    }


	public String getConsSortCode() {
		return consSortCode;
	}

	public void setConsSortCode(String consSortCode) {
		this.consSortCode = consSortCode;
	}

	public String getConsNo() {
		return consNo;
	}

	public void setConsNo(String consNo) {
		this.consNo = consNo;
	}

	public String getSearch() {
		return search;
	}

	public void setSearch(String search) {
		this.search = search;
	}

	/**
	 * @return the saleOrder
	 */
	public String getSaleOrder() {
		return saleOrder;
	}

	/**
	 * @param saleOrder the saleOrder to set
	 */
	public void setSaleOrder(String saleOrder) {
		this.saleOrder = saleOrder;
	}

	/**
	 * @return the goodsCode
	 */
	public String getGoodsCode() {
		return goodsCode;
	}

	/**
	 * @param goodsCode the goodsCode to set
	 */
	public void setGoodsCode(String goodsCode) {
		this.goodsCode = goodsCode;
	}

	/**
	 * @return the pay
	 */
	public String getPay() {
		return pay;
	}

	/**
	 * @param pay the pay to set
	 */
	public void setPay(String pay) {
		this.pay = pay;
	}

	/**
	 * @return the orderNo
	 */
	public String getOrderNo() {
		return orderNo;
	}

	/**
	 * @param orderNo the orderNo to set
	 */
	public void setOrderNo(String orderNo) {
		this.orderNo = orderNo;
	}

	/**
	 * @return the yxBackResult
	 */
	public String getYxBackResult() {
		return yxBackResult;
	}

	/**
	 * @param yxBackResult the yxBackResult to set
	 */
	public void setYxBackResult(String yxBackResult) {
		this.yxBackResult = yxBackResult;
	}
}
