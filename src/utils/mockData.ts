export interface MockDataset {
  headers: string[];
  data: Record<string, any>[];
  filename: string;
}

export const getShopifyMockData = (): MockDataset => {
  const data = [
    {
      "Name": "#1024-A",
      "Created at": "2026-05-12 14:23:10 +0530",
      "Email": "aman.sharma@gmail.com",
      "Financial Status": "paid",
      "Fulfillment Status": "fulfilled",
      "Total": 2499,
      "Discount Amount": 250,
      "Refunded Amount": 0,
      "Lineitem name": "Wellness Gummies Sleep-Well",
      "Lineitem quantity": 2,
      "Lineitem price": 1249.5,
      "Payment Method": "prepaid",
      "Shipping Name": "Aman Sharma",
      "Phone": "9876543210",
      "Billing City": "Mumbai",
      "Billing State": "Maharashtra"
    },
    {
      "Name": "#1025-B",
      "Created at": "2026-05-14 10:15:30 +0530",
      "Email": "priya.k@yahoo.com",
      "Financial Status": "pending",
      "Fulfillment Status": "unfulfilled",
      "Total": 1899,
      "Discount Amount": 0,
      "Refunded Amount": 0,
      "Lineitem name": "Anti-Aging Collagen Serum",
      "Lineitem quantity": 1,
      "Lineitem price": 1899,
      "Payment Method": "cod",
      "Shipping Name": "Priya Krishnan",
      "Phone": "9988776655",
      "Billing City": "Bengaluru",
      "Billing State": "Karnataka"
    },
    {
      "Name": "#1026-C",
      "Created at": "2026-05-15 18:45:00 +0530",
      "Email": "rahul.varma@outlook.com",
      "Financial Status": "paid",
      "Fulfillment Status": "fulfilled",
      "Total": 4500,
      "Discount Amount": 500,
      "Refunded Amount": 0,
      "Lineitem name": "Hair Vitality Pack Premium",
      "Lineitem quantity": 1,
      "Lineitem price": 4500,
      "Payment Method": "prepaid",
      "Shipping Name": "Rahul Varma",
      "Phone": "9123456789",
      "Billing City": "New Delhi",
      "Billing State": "Delhi"
    },
    {
      "Name": "#1027-D",
      "Created at": "2026-05-16 09:30:15 +0530",
      "Email": "sneha.patel@gmail.com",
      "Financial Status": "refunded",
      "Fulfillment Status": "fulfilled",
      "Total": 1299,
      "Discount Amount": 100,
      "Refunded Amount": 1299,
      "Lineitem name": "Hydrating Face Cleanser Mint",
      "Lineitem quantity": 1,
      "Lineitem price": 1299,
      "Payment Method": "prepaid",
      "Shipping Name": "Sneha Patel",
      "Phone": "9812345678",
      "Billing City": "Ahmedabad",
      "Billing State": "Gujarat"
    },
    {
      "Name": "#1028-E",
      "Created at": "2026-05-17 11:05:00 +0530",
      "Email": "vikram.singh@rediffmail.com",
      "Financial Status": "paid",
      "Fulfillment Status": "fulfilled",
      "Total": 5999,
      "Discount Amount": 600,
      "Refunded Amount": 0,
      "Lineitem name": "Complete Skin Rejuvenation Kit",
      "Lineitem quantity": 1,
      "Lineitem price": 5999,
      "Payment Method": "prepaid",
      "Shipping Name": "Vikram Singh",
      "Phone": "9555667788",
      "Billing City": "Jaipur",
      "Billing State": "Rajasthan"
    },
    {
      "Name": "#1029-F",
      "Created at": "2026-05-18 16:20:12 +0530",
      "Email": "amit.g@gmail.com",
      "Financial Status": "pending",
      "Fulfillment Status": "unfulfilled",
      "Total": 3198,
      "Discount Amount": 0,
      "Refunded Amount": 0,
      "Lineitem name": "Anti-Aging Collagen Serum",
      "Lineitem quantity": 2,
      "Lineitem price": 1599,
      "Payment Method": "cod",
      "Shipping Name": "Amit Gupta",
      "Phone": "9777888999",
      "Billing City": "Pune",
      "Billing State": "Maharashtra"
    },
    {
      "Name": "#1030-G",
      "Created at": "2026-05-19 12:40:00 +0530",
      "Email": "ananya.d@gmail.com",
      "Financial Status": "paid",
      "Fulfillment Status": "fulfilled",
      "Total": 2498,
      "Discount Amount": 200,
      "Refunded Amount": 0,
      "Lineitem name": "Wellness Gummies Sleep-Well",
      "Lineitem quantity": 2,
      "Lineitem price": 1249,
      "Payment Method": "prepaid",
      "Shipping Name": "Ananya Das",
      "Phone": "9666555444",
      "Billing City": "Kolkata",
      "Billing State": "West Bengal"
    },
    {
      "Name": "#1031-H",
      "Created at": "2026-05-20 19:10:45 +0530",
      "Email": "karan.j@outlook.com",
      "Financial Status": "paid",
      "Fulfillment Status": "fulfilled",
      "Total": 1899,
      "Discount Amount": 0,
      "Refunded Amount": 0,
      "Lineitem name": "Anti-Aging Collagen Serum",
      "Lineitem quantity": 1,
      "Lineitem price": 1899,
      "Payment Method": "prepaid",
      "Shipping Name": "Karan Johar",
      "Phone": "9111222333",
      "Billing City": "Mumbai",
      "Billing State": "Maharashtra"
    }
  ];

  return {
    headers: Object.keys(data[0]),
    data,
    filename: "Shopify_Live_Sales_Mock.xlsx"
  };
};

export const getLogisticsMockData = (): MockDataset => {
  const data = [
    {
      "Order ID": "SR-9080",
      "Channel SKU": "MED-SLEEP-GUM",
      "Product Name": "Wellness Gummies Sleep-Well",
      "Product Quantity": 2,
      "Status": "DELIVERED",
      "AWB Code": "7864098231",
      "Courier Company": "Delhivery",
      "Pickup Address Name": "Mumbai Warehouse",
      "Zone": "z_a",
      "Address State": "Maharashtra",
      "Address City": "Pune",
      "Customer Name": "Ajay Mehta",
      "Order Total": 2499,
      "COD Payble Amount": 0,
      "Freight Total Amount": 120
    },
    {
      "Order ID": "SR-9081",
      "Channel SKU": "SKN-COLLAGEN",
      "Product Name": "Anti-Aging Collagen Serum",
      "Product Quantity": 1,
      "Status": "RTO DELIVERED",
      "AWB Code": "5432109876",
      "Courier Company": "Xpressbees",
      "Pickup Address Name": "Mumbai Warehouse",
      "Zone": "z_c",
      "Address State": "Uttar Pradesh",
      "Address City": "Lucknow",
      "Customer Name": "Suresh Pal",
      "Order Total": 1899,
      "COD Payble Amount": 1899,
      "Freight Total Amount": 185
    },
    {
      "Order ID": "SR-9082",
      "Channel SKU": "MED-HAIR-PACK",
      "Product Name": "Hair Vitality Pack Premium",
      "Product Quantity": 1,
      "Status": "DELIVERED",
      "AWB Code": "9812763450",
      "Courier Company": "Bluedart",
      "Pickup Address Name": "Delhi Hub",
      "Zone": "z_b",
      "Address State": "Haryana",
      "Address City": "Gurugram",
      "Customer Name": "Neha Gupta",
      "Order Total": 4500,
      "COD Payble Amount": 0,
      "Freight Total Amount": 210
    },
    {
      "Order ID": "SR-9083",
      "Channel SKU": "SKN-CLEANSER",
      "Product Name": "Hydrating Face Cleanser Mint",
      "Product Quantity": 1,
      "Status": "CANCELED",
      "AWB Code": "2345678901",
      "Courier Company": "Delhivery",
      "Pickup Address Name": "Delhi Hub",
      "Zone": "z_d",
      "Address State": "Bihar",
      "Address City": "Patna",
      "Customer Name": "Ravi Jha",
      "Order Total": 1299,
      "COD Payble Amount": 1299,
      "Freight Total Amount": 0
    },
    {
      "Order ID": "SR-9084",
      "Channel SKU": "SKN-COLLAGEN",
      "Product Name": "Anti-Aging Collagen Serum",
      "Product Quantity": 2,
      "Status": "DELIVERED",
      "AWB Code": "8877665544",
      "Courier Company": "Delhivery",
      "Pickup Address Name": "Mumbai Warehouse",
      "Zone": "z_b",
      "Address State": "Karnataka",
      "Address City": "Bengaluru",
      "Customer Name": "Anita Roy",
      "Order Total": 3198,
      "COD Payble Amount": 0,
      "Freight Total Amount": 140
    },
    {
      "Order ID": "SR-9085",
      "Channel SKU": "MED-SLEEP-GUM",
      "Product Name": "Wellness Gummies Sleep-Well",
      "Product Quantity": 2,
      "Status": "RTO DELIVERED",
      "AWB Code": "4433221100",
      "Courier Company": "Xpressbees",
      "Pickup Address Name": "Mumbai Warehouse",
      "Zone": "z_d",
      "Address State": "West Bengal",
      "Address City": "Kolkata",
      "Customer Name": "Bipul Das",
      "Order Total": 2498,
      "COD Payble Amount": 2498,
      "Freight Total Amount": 190
    },
    {
      "Order ID": "SR-9086",
      "Channel SKU": "SKN-COLLAGEN",
      "Product Name": "Anti-Aging Collagen Serum",
      "Product Quantity": 1,
      "Status": "DELIVERED",
      "AWB Code": "1122334455",
      "Courier Company": "Bluedart",
      "Pickup Address Name": "Mumbai Warehouse",
      "Zone": "z_a",
      "Address State": "Maharashtra",
      "Address City": "Mumbai",
      "Customer Name": "Rajesh K",
      "Order Total": 1899,
      "COD Payble Amount": 0,
      "Freight Total Amount": 95
    }
  ];

  return {
    headers: Object.keys(data[0]),
    data,
    filename: "Shiprocket_Logistics_Mock.xlsx"
  };
};

export const getUniversalMockData = (): MockDataset => {
  const data = [
    {
      "Date": "2026-05-01",
      "Reference": "REF-8001",
      "Description": "Shopify Hosting Plan",
      "Category": "SaaS Subscriptions",
      "Debit (Expense)": 2999,
      "Credit (Revenue)": 0,
      "Net Balance": -2999,
      "Department": "Operations"
    },
    {
      "Date": "2026-05-02",
      "Reference": "REV-5012",
      "Description": "Batch Settlement Prepaid",
      "Category": "Sales Revenue",
      "Debit (Expense)": 0,
      "Credit (Revenue)": 145000,
      "Net Balance": 145000,
      "Department": "Finance"
    },
    {
      "Date": "2026-05-03",
      "Reference": "REF-8002",
      "Description": "Facebook Retargeting Ads",
      "Category": "Marketing & Ads",
      "Debit (Expense)": 45000,
      "Credit (Revenue)": 0,
      "Net Balance": -45000,
      "Department": "Growth Marketing"
    },
    {
      "Date": "2026-05-04",
      "Reference": "REF-8003",
      "Description": "Delhivery Shipping Freight Charges",
      "Category": "Logistics Freight",
      "Debit (Expense)": 18450,
      "Credit (Revenue)": 0,
      "Net Balance": -18450,
      "Department": "Supply Chain"
    },
    {
      "Date": "2026-05-05",
      "Reference": "REV-5013",
      "Description": "Razorpay COD Daily Remittance",
      "Category": "Sales Revenue",
      "Debit (Expense)": 0,
      "Credit (Revenue)": 89000,
      "Net Balance": 89000,
      "Department": "Finance"
    },
    {
      "Date": "2026-05-06",
      "Reference": "REF-8004",
      "Description": "Supabase Cloud DB Premium",
      "Category": "SaaS Subscriptions",
      "Debit (Expense)": 2100,
      "Credit (Revenue)": 0,
      "Net Balance": -2100,
      "Department": "Engineering"
    },
    {
      "Date": "2026-05-07",
      "Reference": "REF-8005",
      "Description": "Avery AI API Usage Retainer",
      "Category": "SaaS Subscriptions",
      "Debit (Expense)": 12000,
      "Credit (Revenue)": 0,
      "Net Balance": -12000,
      "Department": "Engineering"
    }
  ];

  return {
    headers: Object.keys(data[0]),
    data,
    filename: "SaaS_Company_Ledger_Mock.xlsx"
  };
};
