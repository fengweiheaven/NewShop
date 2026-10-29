window.OrderSeed = {
  "products": [
    {
      "id": "p-30101020200001",
      "internalSku": "OF-001",
      "supplierId": "youben",
      "supplierSku": "30101020200001",
      "name": "OMG我迷家营养蛋 40枚/盒",
      "spec": "40枚/盒",
      "unit": "盒",
      "note": "",
      "aliases": []
    },
    {
      "id": "p-30101120200015",
      "internalSku": "OF-002",
      "supplierId": "youben",
      "supplierSku": "30101120200015",
      "name": "圣迪乐多维谷物蛋 30枚/盒",
      "spec": "30枚/盒",
      "unit": "盒",
      "note": "",
      "aliases": []
    },
    {
      "id": "p-30101120200002",
      "internalSku": "OF-003",
      "supplierId": "youben",
      "supplierSku": "30101120200002",
      "name": "圣迪乐鲜蛋谷物系列 30枚/盒",
      "spec": "30枚/盒",
      "unit": "盒",
      "note": "",
      "aliases": []
    },
    {
      "id": "p-30101120200016",
      "internalSku": "OF-004",
      "supplierId": "youben",
      "supplierSku": "30101120200016",
      "name": "圣迪乐减胆固醇营养蛋30枚/盒电商装",
      "spec": "30枚/盒",
      "unit": "盒",
      "note": "",
      "aliases": []
    },
    {
      "id": "p-30101120200017",
      "internalSku": "OF-005",
      "supplierId": "youben",
      "supplierSku": "30101120200017",
      "name": "圣迪乐减胆固醇营养蛋20枚/盒电商装",
      "spec": "20枚/盒",
      "unit": "盒",
      "note": "",
      "aliases": []
    },
    {
      "id": "p-30101010500045",
      "internalSku": "OF-006",
      "supplierId": "youben",
      "supplierSku": "30101010500045",
      "name": "圣迪乐OMEGA-3营养蛋30枚（儿童）/盒电商装",
      "spec": "30枚/盒",
      "unit": "盒",
      "note": "原儿童营养蛋",
      "aliases": []
    },
    {
      "id": "p-30101010500043",
      "internalSku": "OF-007",
      "supplierId": "youben",
      "supplierSku": "30101010500043",
      "name": "圣迪乐OMEGA-3营养蛋20枚（儿童）/盒电商装",
      "spec": "20枚/盒",
      "unit": "盒",
      "note": "原儿童营养蛋",
      "aliases": []
    },
    {
      "id": "p-30101010500046",
      "internalSku": "OF-008",
      "supplierId": "youben",
      "supplierSku": "30101010500046",
      "name": "圣迪乐OMEGA-3营养蛋30枚/盒电商装",
      "spec": "30枚/盒",
      "unit": "盒",
      "note": "原超级营养蛋",
      "aliases": []
    },
    {
      "id": "p-30101120600010",
      "internalSku": "OF-009",
      "supplierId": "youben",
      "supplierSku": "30101120600010",
      "name": "圣迪乐OMEGA-3营养蛋20枚/盒电商装",
      "spec": "20枚/盒",
      "unit": "盒",
      "note": "",
      "aliases": []
    },
    {
      "id": "p-30201030200023",
      "internalSku": "OF-010",
      "supplierId": "youben",
      "supplierSku": "30201030200023",
      "name": "美玉子日式溏心蛋12枚/盒 无外箱",
      "spec": "12枚/盒",
      "unit": "盒",
      "note": "",
      "aliases": []
    },
    {
      "id": "p-30201030200003",
      "internalSku": "OF-011",
      "supplierId": "youben",
      "supplierSku": "30201030200003",
      "name": "美玉子日式溏心蛋24枚/盒礼品箱",
      "spec": "24枚/盒",
      "unit": "盒",
      "note": "",
      "aliases": []
    }
  ],
  "candidates": [
    {
      "customer": "兮迪",
      "name": "圣迪乐Omega3儿童超级营养蛋（30枚）商品规格:30枚/盒",
      "spec": "30枚/盒",
      "productId": "p-30101010500045",
      "reason": "儿童+OMEGA3+30枚一致；供应商备注原儿童营养蛋；不得误配非儿童30101010500046"
    },
    {
      "customer": "兮迪",
      "name": "圣迪乐可生食减胆固醇营养蛋（30枚）商品规格:30枚/盒",
      "spec": "30枚/盒",
      "productId": "p-30101120200016",
      "reason": "减胆固醇+30枚一致"
    },
    {
      "customer": "GBH",
      "name": "圣迪乐 减胆固醇营养蛋30枚/盒",
      "spec": "30枚/盒",
      "productId": "p-30101120200016",
      "reason": "减胆固醇+30枚一致；上游编码与客户SKU不同",
      "sourceSku": "4300001651001"
    }
  ],
  "supplier": {
    "id": "youben",
    "name": "友本",
    "contact": "",
    "phone": "",
    "note": "名称来自导单模板，可在此完善供应商主体信息",
    "template": {
      "name": "导单模版-友本.xlsx",
      "sheetName": "Sheet1",
      "columns": [
        {
          "header": "订单号",
          "field": "orderNo",
          "required": true,
          "fallback": ""
        },
        {
          "header": "买家昵称",
          "field": "buyerNickname",
          "required": false,
          "fallback": ""
        },
        {
          "header": "收货人",
          "field": "recipient",
          "required": true,
          "fallback": ""
        },
        {
          "header": "详细地址",
          "field": "address",
          "required": true,
          "fallback": ""
        },
        {
          "header": "收货人电话",
          "field": "phone",
          "required": true,
          "fallback": ""
        },
        {
          "header": "商品编码",
          "field": "supplierSku",
          "required": true,
          "fallback": ""
        },
        {
          "header": "数量",
          "field": "quantity",
          "required": true,
          "fallback": ""
        }
      ]
    }
  }
};
