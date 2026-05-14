import { useState, useRef, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell, AreaChart, Area,
  RadarChart, Radar, PolarGrid, PolarAngleAxis,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer
} from "recharts";
// ─── Analytics helpers ───────────────────────────────────────
const aSum=(arr,k)=>arr.reduce((s,r)=>s+Number(r[k]||0),0);
const aAvg=(arr,k)=>arr.length?+(aSum(arr,k)/arr.length).toFixed(1):0;
const aPct=(n,d)=>d?Math.round(n/d*100):0;
const aMonth=(d)=>{try{return["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][new Date(d).getMonth()]||"?";}catch(e){return"?";}};
const GRADES_LIST=["OPC","PPC","ProMaxX"];
const PLANTS_LIST=["Aligarh Plant","Morak Plant","Depot"];
const PAY_LIST=["Advance","Credit","BG","LC"];
const ANALYTICS_DASHBOARDS=[
  {key:"exec",    label:"Executive Overview", icon:"🏠",desc:"KPIs · Volume · Revenue"},
  {key:"workflow",label:"Approval Workflow",  icon:"🔄",desc:"Funnel · TAT · Stage"},
  {key:"ncr",     label:"NCR Analytics",      icon:"📊",desc:"Margin · Grade · Plant"},
  {key:"perf",    label:"Team Performance",   icon:"👥",desc:"SO · ASM · Leaderboard"},
  {key:"sales",   label:"Sales vs Approval",  icon:"💰",desc:"Utilization · Gap analysis"},
  {key:"customer",label:"Customer Intel",     icon:"🏢",desc:"Risk · Volume · Discount"},
  {key:"geo",     label:"Geographic",         icon:"🗺", desc:"Region · Zone analysis"},
  {key:"freight", label:"Freight & Cost",     icon:"🚛",desc:"Cost breakdown · Leakage"},
  {key:"aging",   label:"Approval Aging",     icon:"⏱", desc:"Buckets · Escalation"},
  {key:"audit",   label:"Audit & Compliance", icon:"🔍",desc:"Exceptions · Policy"},
  {key:"trade",   label:"Trade Price",        icon:"💹",desc:"Discount · Competitiveness"},
  {key:"master",  label:"Master Data",        icon:"📦",desc:"Customer · Plant stats"},
  {key:"forecast",label:"Forecast & Plan",    icon:"📅",desc:"Pipeline · Dispatch plan"},
];

// ═══ MASTER DATA ═══════════════════════════════════════════
const INITIAL_CUSTOMERS = [
  {id:1,accountGroup:"Z001",accountGroupDesc:"Trade Customer",status:"Active",division:"10",distributionChannel:"10",salesOfficeCode:"SO01",salesOfficeName:"Jaipur Sales Office",customerCode:"C001",customerName:"Raj Construction Co.",proprietorDirector:"Rajesh Kumar Sharma",co:"",houseNo:"12",street1:"Main Market Road",street2:"Near Bus Stand",street3:"",city:"Jaipur",districtCityCode:"JP001",districtDesc:"Jaipur",tehsilCode:"TH001",tehsilDesc:"Jaipur Central",regionCode:"RJ",regionDesc:"Rajasthan",clusterCode:"CL01",clusterDesc:"North Rajasthan",transpZoneCd:"TZ01",transportationZone:"Zone A",pinCode:"302001",landlineNo:"0141-2345678",emailId:"raj@email.com",mobileNo:"9876543210",plant:"Aligarh Plant",gstin:"08AABCR1234A1Z5",panNumber:"AABCR1234A",taxClass:"TC01",taxClassDesc:"Regular",tanNumber:"JPYR12345D",lstNumber:"LST001",cstNumber:"CST001",tinNumber:"TIN001",dateOfAssociation:"2018-04-01",securityDeposit:"50000",creditLimit:"200000",previousAccountNo:"OLD-C001",customerBlock:"No",incoterm:"CIF",exciseNumber:"ECC001",distanceFromFactory:"120",distributorCode:"D001",distributorName:"North Rajasthan Distributors",retailer:"Yes",consignee:"Self",dealerCode:"DL001",dealerName:"Raj Construction Co.",salesPromoter:"SP001",salesPromoterName:"Amit Verma",salesAreaBlock:"No",bankIFSC:"SBIN0001234",bankAccountNo:"12345678901",bankName:"State Bank of India",accountHolderName:"Rajesh Kumar Sharma",bankBranchName:"Jaipur Main Branch",bankCity:"Jaipur",bankHouseNo:"1",bankStreet:"MI Road",bankState:"Rajasthan",bankPIN:"302001",vidNumber:"VID001",assignedTo:["SO001","ASM001","RH001"]},
  {id:2,accountGroup:"Z002",accountGroupDesc:"Non-Trade Customer",status:"Active",division:"10",distributionChannel:"20",salesOfficeCode:"SO02",salesOfficeName:"Kota Sales Office",customerCode:"C002",customerName:"Sharma Builders Pvt. Ltd.",proprietorDirector:"Deepak Sharma",co:"C/O ABC Holdings",houseNo:"45",street1:"Industrial Area",street2:"Sector 5",street3:"",city:"Kota",districtCityCode:"KT001",districtDesc:"Kota",tehsilCode:"TH002",tehsilDesc:"Kota Urban",regionCode:"RJ",regionDesc:"Rajasthan",clusterCode:"CL02",clusterDesc:"South Rajasthan",transpZoneCd:"TZ02",transportationZone:"Zone B",pinCode:"324001",landlineNo:"0744-2567890",emailId:"sharma@email.com",mobileNo:"9876543211",plant:"Morak Plant",gstin:"08AABCS5678B2Z6",panNumber:"AABCS5678B",taxClass:"TC01",taxClassDesc:"Regular",tanNumber:"KOTR23456E",lstNumber:"LST002",cstNumber:"CST002",tinNumber:"TIN002",dateOfAssociation:"2019-07-15",securityDeposit:"100000",creditLimit:"500000",previousAccountNo:"",customerBlock:"No",incoterm:"FOB",exciseNumber:"ECC002",distanceFromFactory:"45",distributorCode:"D002",distributorName:"South Rajasthan Distributors",retailer:"No",consignee:"Sharma Builders",dealerCode:"DL002",dealerName:"Sharma Builders",salesPromoter:"SP002",salesPromoterName:"Priya Singh",salesAreaBlock:"No",bankIFSC:"HDFC0002345",bankAccountNo:"23456789012",bankName:"HDFC Bank",accountHolderName:"Deepak Sharma",bankBranchName:"Kota Branch",bankCity:"Kota",bankHouseNo:"5",bankStreet:"Station Road",bankState:"Rajasthan",bankPIN:"324001",vidNumber:"VID002",assignedTo:["SO001","ASM001","RH002"]},
  {id:3,accountGroup:"Z001",accountGroupDesc:"Trade Customer",status:"Active",division:"10",distributionChannel:"10",salesOfficeCode:"SO03",salesOfficeName:"Agra Sales Office",customerCode:"C003",customerName:"Gupta Infra Solutions",proprietorDirector:"Ramesh Gupta",co:"",houseNo:"8",street1:"MG Road",street2:"",street3:"",city:"Agra",districtCityCode:"AG001",districtDesc:"Agra",tehsilCode:"TH003",tehsilDesc:"Agra Urban",regionCode:"UP",regionDesc:"Uttar Pradesh",clusterCode:"CL03",clusterDesc:"West UP",transpZoneCd:"TZ03",transportationZone:"Zone C",pinCode:"282001",landlineNo:"0562-2345678",emailId:"gupta@email.com",mobileNo:"9876543212",plant:"Aligarh Plant",gstin:"09AABCG9012C3Z7",panNumber:"AABCG9012C",taxClass:"TC01",taxClassDesc:"Regular",tanNumber:"AGRR34567F",lstNumber:"LST003",cstNumber:"CST003",tinNumber:"TIN003",dateOfAssociation:"2020-01-10",securityDeposit:"75000",creditLimit:"300000",previousAccountNo:"",customerBlock:"No",incoterm:"EXW",exciseNumber:"ECC003",distanceFromFactory:"60",distributorCode:"D003",distributorName:"West UP Distributors",retailer:"Yes",consignee:"Self",dealerCode:"DL003",dealerName:"Gupta Infra Solutions",salesPromoter:"SP003",salesPromoterName:"Suresh Yadav",salesAreaBlock:"No",bankIFSC:"ICIC0003456",bankAccountNo:"34567890123",bankName:"ICICI Bank",accountHolderName:"Ramesh Gupta",bankBranchName:"Agra Branch",bankCity:"Agra",bankHouseNo:"12",bankStreet:"Fatehabad Road",bankState:"Uttar Pradesh",bankPIN:"282001",vidNumber:"VID003",assignedTo:["SO001","ASM002","RH001"]},
  {id:4,accountGroup:"Z003",accountGroupDesc:"Dealer",status:"Inactive",division:"10",distributionChannel:"30",salesOfficeCode:"SO02",salesOfficeName:"Kota Sales Office",customerCode:"C004",customerName:"Mehta Projects Ltd.",proprietorDirector:"Vijay Mehta",co:"",houseNo:"22",street1:"Ring Road",street2:"Near Airport",street3:"",city:"Udaipur",districtCityCode:"UD001",districtDesc:"Udaipur",tehsilCode:"TH004",tehsilDesc:"Udaipur Urban",regionCode:"RJ",regionDesc:"Rajasthan",clusterCode:"CL02",clusterDesc:"South Rajasthan",transpZoneCd:"TZ04",transportationZone:"Zone D",pinCode:"313001",landlineNo:"0294-2345678",emailId:"mehta@email.com",mobileNo:"9876543213",plant:"Morak Plant",gstin:"08AABCM3456D4Z8",panNumber:"AABCM3456D",taxClass:"TC02",taxClassDesc:"Composition",tanNumber:"UDPR45678G",lstNumber:"LST004",cstNumber:"CST004",tinNumber:"TIN004",dateOfAssociation:"2017-06-01",securityDeposit:"150000",creditLimit:"750000",previousAccountNo:"OLD-C004",customerBlock:"Yes",incoterm:"DDP",exciseNumber:"ECC004",distanceFromFactory:"280",distributorCode:"D002",distributorName:"South Rajasthan Distributors",retailer:"No",consignee:"Mehta Projects",dealerCode:"DL004",dealerName:"Mehta Projects Ltd.",salesPromoter:"SP002",salesPromoterName:"Priya Singh",salesAreaBlock:"No",bankIFSC:"PUNB0004567",bankAccountNo:"45678901234",bankName:"Punjab National Bank",accountHolderName:"Vijay Mehta",bankBranchName:"Udaipur Branch",bankCity:"Udaipur",bankHouseNo:"3",bankStreet:"Court Circle",bankState:"Rajasthan",bankPIN:"313001",vidNumber:"VID004",assignedTo:["ASM002","RH003"]},
  {id:5,accountGroup:"Z001",accountGroupDesc:"Trade Customer",status:"Active",division:"10",distributionChannel:"10",salesOfficeCode:"SO01",salesOfficeName:"Jaipur Sales Office",customerCode:"C005",customerName:"Patel Constructions",proprietorDirector:"Hitesh Patel",co:"",houseNo:"55",street1:"Tonk Road",street2:"",street3:"",city:"Jaipur",districtCityCode:"JP001",districtDesc:"Jaipur",tehsilCode:"TH001",tehsilDesc:"Jaipur Central",regionCode:"RJ",regionDesc:"Rajasthan",clusterCode:"CL01",clusterDesc:"North Rajasthan",transpZoneCd:"TZ01",transportationZone:"Zone A",pinCode:"302015",landlineNo:"",emailId:"patel@email.com",mobileNo:"9876543214",plant:"Aligarh Plant",gstin:"08AABCP7890E5Z9",panNumber:"AABCP7890E",taxClass:"TC01",taxClassDesc:"Regular",tanNumber:"JPYP56789H",lstNumber:"LST005",cstNumber:"CST005",tinNumber:"TIN005",dateOfAssociation:"2021-03-20",securityDeposit:"25000",creditLimit:"100000",previousAccountNo:"",customerBlock:"No",incoterm:"CIF",exciseNumber:"ECC005",distanceFromFactory:"125",distributorCode:"D001",distributorName:"North Rajasthan Distributors",retailer:"Yes",consignee:"Self",dealerCode:"DL005",dealerName:"Patel Constructions",salesPromoter:"SP001",salesPromoterName:"Amit Verma",salesAreaBlock:"No",bankIFSC:"SBIN0005678",bankAccountNo:"56789012345",bankName:"State Bank of India",accountHolderName:"Hitesh Patel",bankBranchName:"Tonk Road Branch",bankCity:"Jaipur",bankHouseNo:"90",bankStreet:"Tonk Road",bankState:"Rajasthan",bankPIN:"302015",vidNumber:"VID005",assignedTo:["SO001","ASM001","RH001"]},
];
const INITIAL_DISTRICTS=[{id:"D001",name:"Jaipur",primaryFreight:320,secondaryFreight:80},{id:"D002",name:"Kota",primaryFreight:280,secondaryFreight:60},{id:"D003",name:"Ajmer",primaryFreight:300,secondaryFreight:70},{id:"D004",name:"Udaipur",primaryFreight:350,secondaryFreight:90},{id:"D005",name:"Jodhpur",primaryFreight:360,secondaryFreight:95},{id:"D006",name:"Agra",primaryFreight:200,secondaryFreight:50},{id:"D007",name:"Mathura",primaryFreight:180,secondaryFreight:45}];
const INITIAL_TRADE_PRICES={OPC:{Bag:380,Loose:340},PPC:{Bag:360,Loose:320},ProMaxX:{Bag:420,Loose:380}};
const INITIAL_COP_MASTER = [
  {id:"COP001",plant:"Aligarh Plant",grade:"PPC",materialType:"Bag",costPMT:3600,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
  {id:"COP002",plant:"Aligarh Plant",grade:"OPC",materialType:"Bag",costPMT:3700,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
  {id:"COP003",plant:"Morak Plant",  grade:"PPC",materialType:"Bag",costPMT:3500,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
  {id:"COP004",plant:"Depot",        grade:"PPC",materialType:"Bag",costPMT:3800,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
];
const INITIAL_PACKING_MASTER = [
  {id:"PK001",materialType:"Bag",  packingType:"PP Bag",costPerMT:162.56,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
  {id:"PK002",materialType:"Loose",packingType:"Bulk",  costPerMT:0,     effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
];
const INITIAL_PLANT_MASTER = [
  {id:"PL001",code:"MGU1",name:"Aligarh Plant",type:"Plant",state:"Uttar Pradesh",region:"NZ",address:"Aligarh, UP",capacity:"",status:"Active",remarks:""},
  {id:"PL002",code:"MCP1",name:"Morak Plant",  type:"Plant",state:"Rajasthan",    region:"RAJ",address:"Morak, Kota RJ",capacity:"",status:"Active",remarks:""},
  {id:"PL003",code:"DEP1",name:"Jaipur Depot", type:"Depot",state:"Rajasthan",    region:"RAJ",address:"Jaipur, RJ",capacity:"",status:"Active",remarks:""},
  {id:"PL004",code:"DEP2",name:"Agra Depot",   type:"Depot",state:"Uttar Pradesh",region:"NZ",address:"Agra, UP",capacity:"",status:"Active",remarks:""},
];
// Legacy lookups (computed from masters for backward compat)
const COP={"Aligarh Plant":3600,"Morak Plant":3500,"Depot":3800};
const PACKING={Bag:162.56,Loose:0};
const INITIAL_SALES_ENTRIES = [];
// Sale entry: { id, requestId, billingDocNo, billingDocDate, materialGrade, invoiceQty, invoiceValue, ratePerBag, enteredBy, enteredAt }

// All pages that can be access-controlled per user
const PAGE_ACCESS_ITEMS = [
  { key:"dashboard",       label:"Dashboard",        group:"Main",       icon:"📊" },
  { key:"new-request",     label:"New Request",       group:"Main",       icon:"➕" },
  { key:"requests",        label:"Price Requests",    group:"Main",       icon:"📋" },
  { key:"sale-updation",   label:"Sale Updation",     group:"Main",       icon:"📝" },
  { key:"reports",         label:"Price Report",      group:"Analytics",  icon:"📈" },
  { key:"sales-report",    label:"Sales Report",      group:"Analytics",  icon:"📊" },
  { key:"ncr-calculator",  label:"NCR Calculator",    group:"Analytics",  icon:"🧮" },
  { key:"audit",           label:"Audit Log",         group:"Analytics",  icon:"📜" },
  { key:"customer-master", label:"Customer Master",   group:"Admin",      icon:"👥" },
  { key:"admin-masters",   label:"Admin Masters",     group:"Admin",      icon:"⚙️" },
];

const INITIAL_TPC_AGENTS=[{id:"TPC001",name:"Agarwal Trading Co."},{id:"TPC002",name:"Singh Enterprises"},{id:"TPC003",name:"Kumar & Associates"}];
const INITIAL_MODE_MASTER=[{id:"MOD001",name:"Road"},{id:"MOD002",name:"Rail"},{id:"MOD003",name:"Pipeline"},{id:"MOD004",name:"Waterway"}];
const INITIAL_UNIT_SOURCE_MASTER=[{id:"US001",name:"MT"},{id:"US002",name:"Bag"},{id:"US003",name:"KG"},{id:"US004",name:"Tonne"}];
const INITIAL_SOURCE_MASTER=[{id:"SRC001",name:"Factory Direct"},{id:"SRC002",name:"Depot"},{id:"SRC003",name:"Third Party"},{id:"SRC004",name:"Import"}];
const INITIAL_STORAGE_LOCATION_MASTER=[{id:"SL001",name:"Warehouse A"},{id:"SL002",name:"Warehouse B"},{id:"SL003",name:"Open Yard"},{id:"SL004",name:"Bonded Warehouse"}];
// OP Commission Rate Master: grade × paymentTerms → rate (₹/MT)
const INITIAL_OP_COMMISSION_MASTER = [
  {id:"OPC001",grade:"OPC",paymentTerms:"Advance",  commissionRate:5,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
  {id:"OPC002",grade:"OPC",paymentTerms:"Credit",   commissionRate:6,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
  {id:"OPC003",grade:"OPC",paymentTerms:"BG",       commissionRate:4,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
  {id:"OPC004",grade:"OPC",paymentTerms:"LC",       commissionRate:4.5,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
  {id:"OPC005",grade:"PPC",paymentTerms:"Advance",  commissionRate:4,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
  {id:"OPC006",grade:"PPC",paymentTerms:"Credit",   commissionRate:5,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
  {id:"OPC007",grade:"PPC",paymentTerms:"BG",       commissionRate:3.5,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
  {id:"OPC008",grade:"PPC",paymentTerms:"LC",       commissionRate:3.5,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
  {id:"OPC009",grade:"ProMaxX",paymentTerms:"Advance",commissionRate:6,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
  {id:"OPC010",grade:"ProMaxX",paymentTerms:"Credit",commissionRate:7,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
  {id:"OPC011",grade:"ProMaxX",paymentTerms:"BG",   commissionRate:5,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
  {id:"OPC012",grade:"ProMaxX",paymentTerms:"LC",   commissionRate:5,effectiveFrom:"2024-01-01",effectiveTo:"9999-12-31",remarks:""},
];

// Price Form Fields that can be restricted per user
const PRICE_FORM_FIELDS = [
  {key:"tradePrice",  label:"Trade Price (₹)"},
  {key:"difference",  label:"Difference from Trade"},
  {key:"netOfGST",    label:"Net of GST"},
  {key:"opCommission",label:"OP Commission"},
  {key:"ncrPmt",      label:"NCR / MT"},
  {key:"cop",         label:"Cost of Production"},
];

const INITIAL_USERS=[
  {id:"SO001",name:"Amit Sharma",role:"Sales Officer",email:"amit@mangalam.com",password:"amit123",status:"Active",assignedCustomers:["C001","C002","C005"],formFieldRights:[],analyticsRights:[]},
  {id:"ASM001",name:"Rajesh Kumar",role:"Area Sales Manager",email:"rajesh@mangalam.com",password:"rajesh123",status:"Active",assignedCustomers:["C001","C002","C005"],formFieldRights:[],analyticsRights:[]},
  {id:"ASM002",name:"Priya Singh",role:"Area Sales Manager",email:"priya@mangalam.com",password:"priya123",status:"Active",assignedCustomers:["C003","C004"],formFieldRights:[],analyticsRights:[]},
  {id:"RH001",name:"Suresh Mehta",role:"Regional Head",email:"suresh@mangalam.com",password:"suresh123",status:"Active",assignedCustomers:["C001","C003","C005"],formFieldRights:[],analyticsRights:[]},
  {id:"RH002",name:"Deepak Verma",role:"Regional Head",email:"deepak@mangalam.com",password:"deepak123",status:"Active",assignedCustomers:["C002"],formFieldRights:[],analyticsRights:[]},
  {id:"RH003",name:"Kavita Joshi",role:"Regional Head",email:"kavita@mangalam.com",password:"kavita123",status:"Active",assignedCustomers:["C004"],formFieldRights:[],analyticsRights:[]},
  {id:"ZH001",name:"Narendra Gupta",role:"Zonal Head",email:"narendra@mangalam.com",password:"narendra123",status:"Active",assignedCustomers:["C001","C002","C003","C004","C005"],formFieldRights:[],analyticsRights:[]},
  {id:"SA001",name:"Anita Patel",role:"Sales & Accounts",email:"anita@mangalam.com",password:"anita123",status:"Active",assignedCustomers:["C001","C002","C003","C004","C005"],formFieldRights:[],analyticsRights:[]},
  {id:"ADM001",name:"Admin User",role:"Admin",email:"admin@mangalam.com",password:"admin123",status:"Active",assignedCustomers:["C001","C002","C003","C004","C005"],formFieldRights:[],analyticsRights:[]},
];
const INITIAL_REQUESTS=[
  {id:"REQ-2024-001",date:"2024-01-15",customerCode:"C001",customerName:"Raj Construction Co.",region:"Rajasthan",zone:"Zone A",destination:"D001",grade:"OPC",materialType:"Bag",qty:500,orderPrice:350,tradePrice:380,difference:-30,netOfGST:296.61,primaryFreight:320,secondaryFreight:80,demrage:10,stationHandling:5,costOfProduction:210,packing:25,unloadingPrice:8,opCommission:5,spCommission:3,totalExpenses:666,ncrPmt:-316,paymentTerms:"Advance",supplyFrom:"Aligarh Plant",validityFrom:"2024-01-20",validityTo:"2024-02-20",tpcAgent:"TPC001",status:"Pending",currentLevel:"Area Sales Manager",createdBy:"SO001",blocked:false,history:[{level:"Sales Officer",action:"Created",by:"SO001",time:"2024-01-15 10:30",remark:""}]},
  {id:"REQ-2024-002",date:"2024-01-16",customerCode:"C002",customerName:"Sharma Builders Pvt. Ltd.",region:"Rajasthan",zone:"Zone B",destination:"D002",grade:"PPC",materialType:"Bag",qty:300,orderPrice:340,tradePrice:360,difference:-20,netOfGST:288.14,primaryFreight:280,secondaryFreight:60,demrage:0,stationHandling:0,costOfProduction:195,packing:25,unloadingPrice:8,opCommission:4,spCommission:3,totalExpenses:575,ncrPmt:-235,paymentTerms:"Credit",supplyFrom:"Morak Plant",validityFrom:"2024-01-20",validityTo:"2024-02-20",tpcAgent:"TPC002",status:"Approved",currentLevel:"Admin",createdBy:"SO001",blocked:false,history:[{level:"Sales Officer",action:"Created",by:"SO001",time:"2024-01-16 09:00",remark:""},{level:"Area Sales Manager",action:"Validated",by:"ASM001",time:"2024-01-16 11:00",remark:"Good deal"},{level:"Regional Head",action:"Approved",by:"RH002",time:"2024-01-16 14:00",remark:"Approved"},{level:"Zonal Head",action:"Approved",by:"ZH001",time:"2024-01-17 10:00",remark:"OK"},{level:"Sales & Accounts",action:"Approved",by:"SA001",time:"2024-01-17 15:00",remark:"Cleared"}]},
  {id:"REQ-2024-003",date:"2024-01-17",customerCode:"C003",customerName:"Gupta Infra Solutions",region:"Uttar Pradesh",zone:"Zone C",destination:"D006",grade:"ProMaxX",materialType:"Loose",qty:800,orderPrice:360,tradePrice:380,difference:-20,netOfGST:305.08,primaryFreight:200,secondaryFreight:50,demrage:15,stationHandling:8,costOfProduction:240,packing:0,unloadingPrice:0,opCommission:6,spCommission:4,totalExpenses:523,ncrPmt:-183,paymentTerms:"BG",supplyFrom:"Depot",validityFrom:"2024-01-22",validityTo:"2024-02-22",tpcAgent:"TPC003",status:"Rejected",currentLevel:"Regional Head",createdBy:"SO001",blocked:false,history:[{level:"Sales Officer",action:"Created",by:"SO001",time:"2024-01-17 08:00",remark:""},{level:"Area Sales Manager",action:"Validated",by:"ASM002",time:"2024-01-17 10:30",remark:"Reviewed"},{level:"Regional Head",action:"Rejected",by:"RH001",time:"2024-01-17 16:00",remark:"Price too low"}]},
  {id:"REQ-2024-004",date:"2024-01-18",customerCode:"C004",customerName:"Mehta Projects Ltd.",region:"Rajasthan",zone:"Zone D",destination:"D004",grade:"OPC",materialType:"Bag",qty:200,orderPrice:370,tradePrice:380,difference:-10,netOfGST:313.56,primaryFreight:350,secondaryFreight:90,demrage:0,stationHandling:0,costOfProduction:210,packing:25,unloadingPrice:8,opCommission:5,spCommission:3,totalExpenses:691,ncrPmt:-321,paymentTerms:"Advance",supplyFrom:"Aligarh Plant",validityFrom:"2024-01-23",validityTo:"2024-02-23",tpcAgent:"TPC001",status:"Pending",currentLevel:"Regional Head",createdBy:"ASM002",blocked:false,history:[{level:"Area Sales Manager",action:"Created",by:"ASM002",time:"2024-01-18 09:00",remark:""}]},
  {id:"REQ-2024-005",date:"2024-01-19",customerCode:"C005",customerName:"Patel Constructions",region:"Rajasthan",zone:"Zone A",destination:"D001",grade:"PPC",materialType:"Loose",qty:600,orderPrice:330,tradePrice:320,difference:10,netOfGST:279.66,primaryFreight:320,secondaryFreight:80,demrage:5,stationHandling:3,costOfProduction:195,packing:0,unloadingPrice:0,opCommission:4,spCommission:3,totalExpenses:610,ncrPmt:-280,paymentTerms:"Credit",supplyFrom:"Morak Plant",validityFrom:"2024-01-24",validityTo:"2024-02-24",tpcAgent:"TPC002",status:"Approved",currentLevel:"Admin",createdBy:"SO001",blocked:true,blockReason:"Customer credit limit exceeded",blockedAt:"2024-01-25 10:00",blockedBy:"ADM001",history:[{level:"Sales Officer",action:"Created",by:"SO001",time:"2024-01-19 10:00",remark:""},{level:"Area Sales Manager",action:"Validated",by:"ASM001",time:"2024-01-19 12:00",remark:"Looks good"},{level:"Regional Head",action:"Approved",by:"RH001",time:"2024-01-19 15:00",remark:"Approved"}]},
];

// ═══ PURE HELPERS (no hooks) ═══════════════════════════════
let reqCtr=6;
const genId=()=>`REQ-2024-${String(reqCtr++).padStart(3,"0")}`;
const ROLE_FLOW=["Sales Officer","Area Sales Manager","Regional Head","Zonal Head","Sales & Accounts","Admin"];
const canAct=(user,req)=>{if(req.status==="Approved"||req.status==="Completed")return false;if(req.status==="Rejected")return user.role==="Regional Head";return req.currentLevel===user.role;};
// SO can resubmit when the req was sent back to their level
const isSentBack=(req)=>req.status==="Pending"&&req.currentLevel==="Sales Officer"&&(req.history||[]).some(h=>h.action==="Sent Back");
const prevLevel=cur=>{const i=ROLE_FLOW.indexOf(cur);return i>0?ROLE_FLOW[i-1]:ROLE_FLOW[0];};
const nextLevel=cur=>{const i=ROLE_FLOW.indexOf(cur);return i<ROLE_FLOW.length-1?ROLE_FLOW[i+1]:null;};
const TODAY=new Date().toISOString().split("T")[0];
const fmt=n=>(n==null||n===""?0:Number(n));

function esc(v){const s=String(v??"");return s.includes(",")||s.includes('"')||s.includes("\n")?`"${s.replace(/"/g,'""')}"`:s;}
function buildCsv(rows,cols){return[cols.map(c=>c.label).join(","),...rows.map(r=>cols.map(c=>esc(r[c.key])).join(","))].join("\n");}
function buildTemplateCsv(cols){return cols.map(c=>c.label).join(",")+"\n";}
function parseCsvToObjects(text,colDefs){
  const lines=text.trim().split("\n");if(lines.length<2)return[];
  const hdrs=lines[0].split(",").map(h=>h.trim().replace(/^"|"$/g,""));
  const km={};hdrs.forEach((h,i)=>{const c=colDefs.find(c=>c.label.toLowerCase()===h.toLowerCase());if(c)km[i]=c.key;});
  return lines.slice(1).map((ln,li)=>{const vals=ln.split(",");const obj={id:Date.now()+li};colDefs.forEach(c=>{obj[c.key]="";});vals.forEach((v,j)=>{if(km[j])obj[km[j]]=v.trim().replace(/^"|"$/g,"");});return obj;}).filter(r=>Object.values(r).some(v=>v&&v!==""));
}

// ═══ COLUMN DEFINITIONS ════════════════════════════════════
const ALL_CM_COLS=[{key:"customerCode",label:"Customer Code"},{key:"customerName",label:"Customer Name"},{key:"accountGroup",label:"Account Group"},{key:"accountGroupDesc",label:"Acct Group Desc"},{key:"status",label:"Status"},{key:"division",label:"Division"},{key:"distributionChannel",label:"Dist. Channel"},{key:"salesOfficeCode",label:"Sales Office Code"},{key:"salesOfficeName",label:"Sales Office Name"},{key:"proprietorDirector",label:"Proprietor/Director"},{key:"co",label:"C/O"},{key:"houseNo",label:"House No"},{key:"street1",label:"Street 1"},{key:"street2",label:"Street 2"},{key:"street3",label:"Street 3"},{key:"city",label:"City"},{key:"districtCityCode",label:"District Code"},{key:"districtDesc",label:"District"},{key:"tehsilCode",label:"Tehsil Code"},{key:"tehsilDesc",label:"Tehsil"},{key:"regionCode",label:"Region Code"},{key:"regionDesc",label:"Region"},{key:"clusterCode",label:"Cluster Code"},{key:"clusterDesc",label:"Cluster"},{key:"transpZoneCd",label:"Transp Zone Cd"},{key:"transportationZone",label:"Transport Zone"},{key:"pinCode",label:"PIN"},{key:"landlineNo",label:"Landline"},{key:"emailId",label:"Email"},{key:"mobileNo",label:"Mobile"},{key:"plant",label:"Plant"},{key:"gstin",label:"GSTIN"},{key:"panNumber",label:"PAN"},{key:"taxClass",label:"Tax Class"},{key:"taxClassDesc",label:"Tax Class Desc"},{key:"tanNumber",label:"TAN"},{key:"lstNumber",label:"LST No"},{key:"cstNumber",label:"CST No"},{key:"tinNumber",label:"TIN No"},{key:"dateOfAssociation",label:"Date of Association"},{key:"securityDeposit",label:"Security Deposit"},{key:"creditLimit",label:"Credit Limit"},{key:"previousAccountNo",label:"Prev Account No"},{key:"customerBlock",label:"Customer Block"},{key:"incoterm",label:"Incoterm"},{key:"exciseNumber",label:"Excise No (ECC)"},{key:"distanceFromFactory",label:"Distance (km)"},{key:"distributorCode",label:"Distributor Code"},{key:"distributorName",label:"Distributor Name"},{key:"retailer",label:"Retailer"},{key:"consignee",label:"Consignee"},{key:"dealerCode",label:"Dealer Code"},{key:"dealerName",label:"Dealer Name"},{key:"salesPromoter",label:"Sales Promoter"},{key:"salesPromoterName",label:"SP Name"},{key:"salesAreaBlock",label:"Sales Area Block"},{key:"bankIFSC",label:"Bank IFSC"},{key:"bankAccountNo",label:"Bank A/C No"},{key:"bankName",label:"Bank Name"},{key:"accountHolderName",label:"A/C Holder Name"},{key:"bankBranchName",label:"Branch"},{key:"bankCity",label:"Bank City"},{key:"bankHouseNo",label:"Bank House No"},{key:"bankStreet",label:"Bank Street"},{key:"bankState",label:"Bank State"},{key:"bankPIN",label:"Bank PIN"},{key:"vidNumber",label:"VID No"}];
const DISTRICT_COLS=[{key:"id",label:"District ID"},{key:"name",label:"District Name"},{key:"primaryFreight",label:"Primary Freight"},{key:"secondaryFreight",label:"Secondary Freight"}];
const TPC_COLS=[{key:"id",label:"Agent ID"},{key:"name",label:"Agent Name"}];
const USER_COLS=[{key:"id",label:"User ID"},{key:"name",label:"Full Name"},{key:"role",label:"Role"},{key:"email",label:"Email"},{key:"status",label:"Status"},{key:"designation",label:"Designation"}];
const TRADE_PRICE_COLS=[{key:"grade",label:"Grade"},{key:"materialType",label:"Material Type"},{key:"price",label:"Trade Price"}];
const REQUEST_REPORT_COLS=[{key:"id",label:"Request ID"},{key:"date",label:"Date"},{key:"customerCode",label:"Customer Code"},{key:"customerName",label:"Customer Name"},{key:"region",label:"Region"},{key:"zone",label:"Zone"},{key:"grade",label:"Grade"},{key:"materialType",label:"Material Type"},{key:"qty",label:"Qty (MT)"},{key:"orderPrice",label:"Order Price"},{key:"tradePrice",label:"Trade Price"},{key:"difference",label:"Difference"},{key:"netOfGST",label:"Net of GST"},{key:"route1",label:"Route 1 (Primary)"},{key:"primaryFreight",label:"Primary Freight"},{key:"route2",label:"Route 2 (Secondary)"},{key:"secondaryFreight",label:"Secondary Freight"},{key:"totalExpenses",label:"Total Expenses"},{key:"ncrPmt",label:"NCR/MT"},{key:"paymentTerms",label:"Payment Terms"},{key:"supplyFrom",label:"Supply From"},{key:"validityFrom",label:"Validity From"},{key:"validityTo",label:"Validity To"},{key:"mode",label:"Mode"},{key:"unitSource",label:"Unit / Source"},{key:"source",label:"Source"},{key:"storageLocation",label:"Storage Location"},{key:"status",label:"Status"},{key:"currentLevel",label:"Current Level"}];
const DEFAULT_CM_COLS=["customerCode","customerName","mobileNo","emailId","districtDesc","regionDesc","plant","gstin","status","creditLimit"];
const EMPTY_CUST={id:null,accountGroup:"",accountGroupDesc:"",status:"Active",division:"",distributionChannel:"",salesOfficeCode:"",salesOfficeName:"",customerCode:"",customerName:"",proprietorDirector:"",co:"",houseNo:"",street1:"",street2:"",street3:"",city:"",districtCityCode:"",districtDesc:"",tehsilCode:"",tehsilDesc:"",regionCode:"",regionDesc:"",clusterCode:"",clusterDesc:"",transpZoneCd:"",transportationZone:"",pinCode:"",landlineNo:"",emailId:"",mobileNo:"",plant:"",gstin:"",panNumber:"",taxClass:"",taxClassDesc:"",tanNumber:"",lstNumber:"",cstNumber:"",tinNumber:"",dateOfAssociation:"",securityDeposit:"",creditLimit:"",previousAccountNo:"",customerBlock:"No",incoterm:"",exciseNumber:"",distanceFromFactory:"",distributorCode:"",distributorName:"",retailer:"No",consignee:"",dealerCode:"",dealerName:"",salesPromoter:"",salesPromoterName:"",salesAreaBlock:"No",bankIFSC:"",bankAccountNo:"",bankName:"",accountHolderName:"",bankBranchName:"",bankCity:"",bankHouseNo:"",bankStreet:"",bankState:"",bankPIN:"",vidNumber:"",assignedTo:[]};
const CUST_FORM_SECTIONS=[{title:"Basic / SAP Info",icon:"📋",fields:[{key:"accountGroup",label:"Account Group",req:true},{key:"accountGroupDesc",label:"Account Group Desc"},{key:"status",label:"Status",type:"select",options:["Active","Inactive","Blocked"]},{key:"division",label:"Division"},{key:"distributionChannel",label:"Distribution Channel"},{key:"salesOfficeCode",label:"Sales Office Code"},{key:"salesOfficeName",label:"Sales Office Name"}]},{title:"Customer Details",icon:"👤",fields:[{key:"customerCode",label:"Customer Code",req:true},{key:"customerName",label:"Customer Name",req:true},{key:"proprietorDirector",label:"Proprietor / Director"},{key:"plant",label:"Plant",type:"select",options:["Aligarh Plant","Morak Plant","Depot"]},{key:"dateOfAssociation",label:"Date of Association",type:"date"},{key:"previousAccountNo",label:"Previous Account No"},{key:"customerBlock",label:"Customer Block",type:"select",options:["No","Yes"]}]},{title:"Address",icon:"🏠",fields:[{key:"co",label:"C/O"},{key:"houseNo",label:"House No"},{key:"street1",label:"Street 1"},{key:"street2",label:"Street 2"},{key:"street3",label:"Street 3"},{key:"city",label:"City"},{key:"pinCode",label:"PIN Code"}]},{title:"Geographic",icon:"📍",fields:[{key:"districtCityCode",label:"District City Code"},{key:"districtDesc",label:"District Description"},{key:"tehsilCode",label:"Tehsil Code"},{key:"tehsilDesc",label:"Tehsil Description"},{key:"regionCode",label:"Region Code"},{key:"regionDesc",label:"Region Description"},{key:"clusterCode",label:"Cluster Code"},{key:"clusterDesc",label:"Cluster Description"},{key:"transpZoneCd",label:"Transport Zone Code"},{key:"transportationZone",label:"Transportation Zone"}]},{title:"Contact",icon:"📞",fields:[{key:"landlineNo",label:"Landline No"},{key:"mobileNo",label:"Mobile No",req:true},{key:"emailId",label:"Email ID",type:"email"}]},{title:"Tax & Compliance",icon:"📄",fields:[{key:"gstin",label:"GSTIN Number"},{key:"panNumber",label:"PAN Number"},{key:"taxClass",label:"Tax Class"},{key:"taxClassDesc",label:"Tax Class Description"},{key:"tanNumber",label:"TAN Number"},{key:"lstNumber",label:"LST Number"},{key:"cstNumber",label:"CST Number"},{key:"tinNumber",label:"TIN Number"},{key:"exciseNumber",label:"Excise Number (ECC)"},{key:"vidNumber",label:"VID Number"}]},{title:"Financial",icon:"💳",fields:[{key:"securityDeposit",label:"Security Deposit (₹)",type:"number"},{key:"creditLimit",label:"Credit Limit (₹)",type:"number"},{key:"incoterm",label:"Incoterm",type:"select",options:["CIF","FOB","EXW","DDP","CFR"]},{key:"distanceFromFactory",label:"Distance from Factory (km)",type:"number"}]},{title:"Trade Classification",icon:"🏪",fields:[{key:"distributorCode",label:"Distributor Code"},{key:"distributorName",label:"Distributor Name"},{key:"retailer",label:"Retailer",type:"select",options:["Yes","No"]},{key:"consignee",label:"Consignee"},{key:"dealerCode",label:"Dealer Code"},{key:"dealerName",label:"Dealer Name"},{key:"salesPromoter",label:"Sales Promoter Code"},{key:"salesPromoterName",label:"Sales Promoter Name"},{key:"salesAreaBlock",label:"Sales Area Block",type:"select",options:["No","Yes"]}]},{title:"Bank Details",icon:"🏦",fields:[{key:"bankIFSC",label:"Bank IFSC"},{key:"bankAccountNo",label:"Bank A/C Number"},{key:"bankName",label:"Bank Name"},{key:"accountHolderName",label:"Account Holder Name"},{key:"bankBranchName",label:"Bank Branch Name"},{key:"bankCity",label:"Bank City"},{key:"bankHouseNo",label:"Bank House No"},{key:"bankStreet",label:"Bank Street"},{key:"bankState",label:"Bank State"},{key:"bankPIN",label:"Bank PIN"}]},{title:"User Assignment",icon:"🔗",fields:[]}];
const ROLES=["Sales Officer","Area Sales Manager","Regional Head","Zonal Head","Sales & Accounts","Admin"];
const ROLE_BADGE={"Sales Officer":"b-so","Area Sales Manager":"b-asm","Regional Head":"b-rh","Zonal Head":"b-zh","Sales & Accounts":"b-sa","Admin":"b-admin"};

// ═══ STYLES ════════════════════════════════════════════════
const CSS=`
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=Sora:wght@400;500;600;700;800&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --gold:#D4890A;--gold-dk:#A86A06;--gold-lt:#F5B800;--gold-f:#FFF8EC;--gold-100:#FFF0CC;--gold-200:#FFE08A;
  --cream:#FFF5F5;--white:#FFFFFF;--light-red:#FFF5F5;
  --ink:#1A1200;--ink2:#3D2D00;--ink3:#6B5200;
  --muted:#9C8050;--border:#FECACA;--border2:#FEE2E2;
  --grn:#059669;--grn-f:#ECFDF5;--red:#DC2626;--red-f:#FEF2F2;--blu:#2563EB;--blu-f:#EFF6FF;--ora:#EA580C;--ora-f:#FFF7ED;
  --sh1:0 1px 3px rgba(180,130,0,.08),0 1px 2px rgba(180,130,0,.06);
  --sh2:0 4px 12px rgba(180,130,0,.12),0 2px 4px rgba(180,130,0,.08);
  --sh3:0 10px 30px rgba(0,0,0,.12),0 4px 10px rgba(0,0,0,.08);
  --r:10px;--r-sm:7px;--r-lg:14px;--r-xl:20px;--tr:.2s ease;
}
body{font-family:'Plus Jakarta Sans',sans-serif;background:#FFF5F5;color:var(--ink);font-size:14px;line-height:1.55;min-height:100vh}
.hdr{height:64px;background:#fff;border-bottom:2px solid #FECACA;display:flex;align-items:center;padding:0 24px;gap:16px;position:sticky;top:0;z-index:300;box-shadow:0 2px 8px rgba(192,20,46,.08)}
.hdr-logo{width:42px;height:42px;background:linear-gradient(135deg,var(--gold-lt),var(--gold));border-radius:12px;display:flex;align-items:center;justify-content:center;font-family:'Sora',sans-serif;font-weight:800;color:#fff;font-size:16px;flex-shrink:0;box-shadow:0 4px 12px rgba(212,137,10,.35)}
.hdr-brand{display:flex;flex-direction:column}
.hdr-title{font-family:'Sora',sans-serif;font-size:15px;font-weight:700;color:var(--ink);line-height:1.1;letter-spacing:-.3px}
.hdr-sub{font-size:10.5px;color:var(--muted);letter-spacing:.3px;font-weight:500}
.hdr-sep{flex:1}
.hdr-search{display:flex;align-items:center;gap:8px;background:var(--cream);border:1.5px solid var(--border);border-radius:24px;padding:6px 14px;font-size:13px;color:var(--muted);cursor:pointer;transition:all var(--tr)}
.hdr-search:hover{border-color:var(--gold);background:var(--gold-f)}
.notif{width:38px;height:38px;border-radius:10px;background:var(--cream);border:1.5px solid var(--border);display:flex;align-items:center;justify-content:center;cursor:pointer;font-size:16px;position:relative;transition:all var(--tr)}
.notif:hover{background:var(--gold-f);border-color:var(--gold)}
.notif-dot{position:absolute;top:6px;right:6px;width:8px;height:8px;background:var(--red);border-radius:50%;border:2px solid #fff}
.hdr-avatar{width:38px;height:38px;border-radius:10px;background:linear-gradient(135deg,var(--gold-lt),var(--gold));display:flex;align-items:center;justify-content:center;font-weight:700;color:#fff;font-size:13px;cursor:pointer;transition:box-shadow var(--tr)}
.hdr-avatar:hover{box-shadow:0 0 0 3px var(--gold-200)}
.hdr-user{display:flex;flex-direction:column}
.hdr-uname{font-size:13px;font-weight:700;color:var(--ink);line-height:1.1}
.hdr-urole{font-size:10px;color:var(--muted);font-weight:500}
.logout-btn{display:flex;align-items:center;gap:6px;padding:6px 12px;border-radius:8px;border:1.5px solid var(--border);background:transparent;color:var(--muted);font-size:12px;font-family:inherit;font-weight:600;cursor:pointer;transition:all var(--tr)}
.logout-btn:hover{background:var(--red-f);border-color:var(--red);color:var(--red)}
.layout{display:flex;min-height:calc(100vh - 64px)}
.sidebar{width:240px;background:#fff;border-right:2px solid #FECACA;flex-shrink:0;display:flex;flex-direction:column;padding:16px 12px;position:sticky;top:64px;height:calc(100vh - 64px);overflow-y:auto}
.sb-section-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1.2px;color:var(--muted);padding:8px 12px 4px;margin-top:8px}
.sb-item{display:flex;align-items:center;gap:10px;padding:10px 12px;cursor:pointer;font-size:13px;font-weight:600;color:var(--muted);border-radius:10px;transition:all var(--tr);margin-bottom:2px;border:1.5px solid transparent}
.sb-item:hover{background:var(--gold-f);color:var(--gold-dk)}
.sb-item.active{background:linear-gradient(135deg,var(--gold-f),#FFEFC0);color:var(--gold-dk);border-color:var(--gold-200);font-weight:700;box-shadow:var(--sh1)}
.sb-icon-wrap{width:32px;height:32px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:15px;background:var(--cream);flex-shrink:0;transition:all var(--tr)}
.sb-item.active .sb-icon-wrap{background:var(--gold);box-shadow:0 4px 10px rgba(212,137,10,.3)}
.sb-label{flex:1}
.sb-badge{background:var(--red);color:#fff;font-size:10px;font-weight:700;border-radius:12px;padding:2px 8px;min-width:20px;text-align:center}
.sb-bottom{margin-top:auto;padding-top:12px;border-top:1.5px solid var(--border2)}
.main{flex:1;padding:24px;overflow-y:auto;min-width:0;background:#FFF5F5}
.ph{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:24px;gap:14px;flex-wrap:wrap}
.ph-eyebrow{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:var(--gold-dk);margin-bottom:4px}
.ph-title{font-family:'Sora',sans-serif;font-size:22px;font-weight:700;color:var(--ink);line-height:1.15;letter-spacing:-.4px}
.ph-sub{font-size:13px;color:var(--muted);margin-top:3px;font-weight:500}
.ph-actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:14px;margin-bottom:24px}
.kpi{background:#fff;border-radius:var(--r-lg);padding:18px 20px;border:1.5px solid #FECACA;box-shadow:0 2px 8px rgba(192,20,46,.07);cursor:pointer;transition:all var(--tr);position:relative;overflow:hidden}
.kpi:hover{box-shadow:var(--sh2);transform:translateY(-2px);border-color:var(--gold-200)}
.kpi-accent{position:absolute;top:0;left:0;right:0;height:4px;border-radius:4px 4px 0 0}
.kpi-icon{width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:18px;margin-bottom:12px}
.kpi-val{font-family:'Sora',sans-serif;font-size:28px;font-weight:800;color:var(--ink);line-height:1;letter-spacing:-1px}
.kpi-label{font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-top:6px}
.kpi-trend{font-size:11.5px;font-weight:600;margin-top:6px;display:flex;align-items:center;gap:4px}
.kpi-trend.up{color:var(--grn)}.kpi-trend.down{color:var(--red)}.kpi-trend.neutral{color:var(--muted)}
.card{background:#fff;border-radius:var(--r-lg);border:1.5px solid #FECACA;box-shadow:0 2px 8px rgba(192,20,46,.07)}
.card-hdr{padding:16px 20px;border-bottom:1.5px solid var(--border2);display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap}
.card-title{font-family:'Sora',sans-serif;font-size:14.5px;font-weight:700;color:var(--ink);letter-spacing:-.2px}
.card-sub{font-size:12px;color:var(--muted);font-weight:500;margin-top:2px}
.card-body{padding:20px}
.bar-chart{display:flex;flex-direction:column;gap:10px}
.bar-row{display:flex;align-items:center;gap:10px}
.bar-label{font-size:12px;font-weight:600;color:var(--ink3);width:90px;flex-shrink:0;text-align:right}
.bar-track{flex:1;height:28px;background:var(--cream);border-radius:6px;overflow:hidden;border:1px solid var(--border2)}
.bar-fill{height:100%;border-radius:6px;display:flex;align-items:center;padding-left:10px;font-size:11.5px;font-weight:700;color:#fff;transition:width .6s cubic-bezier(.4,0,.2,1)}
.bar-val{font-size:12px;font-weight:700;color:var(--ink3);width:60px;flex-shrink:0}
.feed{display:flex;flex-direction:column}
.feed-item{display:flex;gap:12px;padding:12px 0;border-bottom:1px solid var(--border2)}
.feed-item:last-child{border-bottom:none}
.feed-dot{width:32px;height:32px;border-radius:10px;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:14px}
.feed-title{font-size:13px;font-weight:700;color:var(--ink)}
.feed-sub{font-size:11.5px;color:var(--muted);margin-top:2px;font-weight:500}
.feed-time{font-size:11px;color:var(--muted);margin-top:4px;font-weight:500}
.tw{overflow-x:auto;border-radius:0 0 var(--r-lg) var(--r-lg)}
table{width:100%;border-collapse:collapse;font-size:13px}
thead tr{background:var(--cream)}
th{padding:11px 14px;text-align:left;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--ink3);border-bottom:1.5px solid var(--border);white-space:nowrap}
td{padding:12px 14px;border-bottom:1px solid var(--border2);vertical-align:middle;color:var(--ink)}
tr:last-child td{border-bottom:none}
tbody tr{transition:background var(--tr)}
tbody tr:hover td{background:var(--gold-f)}
.sticky-col{position:sticky;left:0;background:var(--white);z-index:2;border-right:1px solid var(--border2)}
tbody tr:hover .sticky-col{background:var(--gold-f)}
thead .sticky-col{background:var(--cream);z-index:3}
.badge{display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700;letter-spacing:.3px}
.b-pending{background:#FFF8E1;color:#B45309;border:1px solid #FDE68A}
.b-approved{background:var(--grn-f);color:#065F46;border:1px solid #A7F3D0}
.b-rejected{background:var(--red-f);color:#991B1B;border:1px solid #FECACA}
.b-validated{background:var(--blu-f);color:#1E40AF;border:1px solid #BFDBFE}
.b-active{background:var(--grn-f);color:#065F46;border:1px solid #A7F3D0}
.b-inactive{background:var(--red-f);color:#991B1B;border:1px solid #FECACA}
.b-blocked{background:#FFF8E1;color:#B45309;border:1px solid #FDE68A}
.b-grade{background:var(--gold-f);color:var(--gold-dk);border:1px solid var(--gold-200)}
.b-so{background:#F5F3FF;color:#5B21B6;border:1px solid #DDD6FE}
.b-asm{background:var(--blu-f);color:#1E40AF;border:1px solid #BFDBFE}
.b-rh{background:var(--grn-f);color:#065F46;border:1px solid #A7F3D0}
.b-zh{background:var(--gold-f);color:var(--gold-dk);border:1px solid var(--gold-200)}
.b-sa{background:var(--ora-f);color:#9A3412;border:1px solid #FED7AA}
.b-admin{background:var(--red-f);color:#991B1B;border:1px solid #FECACA}
.btn{display:inline-flex;align-items:center;gap:6px;padding:8px 16px;border-radius:9px;font-size:13px;font-weight:700;cursor:pointer;border:none;transition:all var(--tr);font-family:'Plus Jakarta Sans',sans-serif;white-space:nowrap;letter-spacing:-.1px}
.btn:hover{transform:translateY(-1px);box-shadow:var(--sh2)}
.btn:active{transform:none;box-shadow:none}
.btn-primary{background:linear-gradient(135deg,var(--gold-lt),var(--gold));color:#fff;box-shadow:0 4px 12px rgba(212,137,10,.3)}
.btn-primary:hover{background:linear-gradient(135deg,var(--gold),var(--gold-dk));box-shadow:0 6px 20px rgba(212,137,10,.4)}
.btn-secondary{background:var(--white);color:var(--ink3);border:1.5px solid var(--border)}.btn-secondary:hover{background:var(--cream);border-color:var(--gold)}
.btn-success{background:var(--grn);color:#fff;box-shadow:0 4px 12px rgba(5,150,105,.25)}
.btn-danger{background:var(--red);color:#fff;box-shadow:0 4px 12px rgba(220,38,38,.25)}
.btn-warning{background:var(--ora);color:#fff}
.btn-ghost{background:transparent;color:var(--muted);border:1.5px solid transparent}.btn-ghost:hover{background:var(--gold-f);color:var(--gold-dk)}
.btn-outline{background:transparent;color:var(--gold-dk);border:1.5px solid var(--gold)}.btn-outline:hover{background:var(--gold-f)}
.btn-sm{padding:6px 12px;font-size:12px;border-radius:8px}
.btn-xs{padding:4px 10px;font-size:11px;border-radius:6px}
.btn:disabled{opacity:.4;cursor:not-allowed;transform:none;box-shadow:none}
.fsec{margin-bottom:22px}
.fsec-title{font-family:'Sora',sans-serif;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--gold-dk);border-bottom:2px solid var(--gold-100);padding-bottom:6px;margin-bottom:14px;display:flex;align-items:center;gap:8px}
.fgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px}
.fgrid2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.fg{display:flex;flex-direction:column;gap:4px}
.fl{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--ink3)}
.fl-req::after{content:' *';color:var(--red)}
.fc{padding:9px 12px;border:1.5px solid var(--border);border-radius:9px;font-size:13.5px;font-family:'Plus Jakarta Sans',sans-serif;color:var(--ink);background:var(--white);transition:all var(--tr);width:100%;font-weight:500}
.fc:focus{outline:none;border-color:var(--gold);box-shadow:0 0 0 3px rgba(212,137,10,.12);background:var(--gold-f)}
.fc:disabled,.fc[readonly]{background:var(--cream);color:var(--muted);border-color:var(--border2)}
.fh{font-size:11px;color:var(--muted);font-weight:500}
.sf{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:16px}
.si-wrap{position:relative;flex:1;min-width:200px}
.si-ico{position:absolute;left:11px;top:50%;transform:translateY(-50%);color:var(--muted);font-size:14px;pointer-events:none}
.si{width:100%;padding:9px 12px 9px 34px;border:1.5px solid var(--border);border-radius:9px;font-size:13px;font-family:inherit;transition:all var(--tr);background:var(--white);font-weight:500;color:var(--ink)}
.si:focus{outline:none;border-color:var(--gold);box-shadow:0 0 0 3px rgba(212,137,10,.12);background:var(--gold-f)}
.fsel{padding:9px 12px;border:1.5px solid var(--border);border-radius:9px;font-size:13px;font-family:inherit;background:var(--white);color:var(--ink);font-weight:500;transition:all var(--tr)}
.fsel:focus{outline:none;border-color:var(--gold)}
.ncr-box{background:var(--gold-f);border:1.5px solid var(--gold-200);border-radius:var(--r);padding:14px 16px}
.ncr-box.critical{background:var(--red-f);border-color:#FECACA}
.ncr-val{font-family:'Sora',sans-serif;font-size:30px;font-weight:800;letter-spacing:-1px}
.ncr-neg{color:var(--red)}.ncr-pos{color:var(--grn)}
.overlay{position:fixed;inset:0;background:rgba(26,18,0,.55);backdrop-filter:blur(4px);z-index:500;display:flex;align-items:center;justify-content:center;padding:16px;animation:fIn .2s ease}
.modal{background:var(--white);border-radius:var(--r-xl);width:100%;max-width:960px;max-height:92vh;display:flex;flex-direction:column;box-shadow:0 24px 80px rgba(0,0,0,.25);animation:sUp .25s cubic-bezier(.34,1.56,.64,1);border:1.5px solid var(--border)}
.modal-xl{max-width:1080px}.modal-sm{max-width:500px}.modal-xs{max-width:400px}
.m-hdr{padding:18px 22px;display:flex;align-items:center;justify-content:space-between;background:var(--white);border-bottom:1.5px solid var(--border2);border-radius:var(--r-xl) var(--r-xl) 0 0;flex-shrink:0}
.m-hdr-accent{width:36px;height:36px;border-radius:10px;background:linear-gradient(135deg,var(--gold-lt),var(--gold));display:flex;align-items:center;justify-content:center;font-size:16px;margin-right:12px;flex-shrink:0;box-shadow:0 4px 12px rgba(212,137,10,.3)}
.m-title{font-family:'Sora',sans-serif;font-size:17px;font-weight:700;color:var(--ink);letter-spacing:-.3px}
.m-sub{font-size:12px;color:var(--muted);margin-top:2px;font-weight:500}
.m-close{width:32px;height:32px;border-radius:8px;background:var(--cream);border:1.5px solid var(--border);color:var(--muted);font-size:18px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all var(--tr)}
.m-close:hover{background:var(--red-f);border-color:var(--red);color:var(--red)}
.m-body{flex:1;overflow-y:auto;padding:22px}
.m-footer{padding:14px 22px;border-top:1.5px solid var(--border2);display:flex;gap:8px;justify-content:flex-end;flex-shrink:0;background:var(--cream);border-radius:0 0 var(--r-xl) var(--r-xl)}
@keyframes fIn{from{opacity:0}to{opacity:1}}
@keyframes sUp{from{transform:translateY(30px) scale(.96);opacity:0}to{transform:translateY(0) scale(1);opacity:1}}
.tabs{display:flex;gap:4px;margin-bottom:18px;background:var(--cream);border-radius:10px;padding:4px;border:1.5px solid var(--border2)}
.tab{padding:7px 16px;cursor:pointer;font-size:12.5px;font-weight:600;color:var(--muted);border-radius:7px;transition:all var(--tr);white-space:nowrap}
.tab.active{background:var(--white);color:var(--gold-dk);box-shadow:var(--sh1);font-weight:700}
.tab:hover:not(.active){color:var(--ink3);background:var(--gold-f)}
.htl{display:flex;flex-direction:column}
.hi{display:flex;gap:14px;padding:12px 0;position:relative}
.hi:not(:last-child)::after{content:'';position:absolute;left:15px;top:42px;bottom:0;width:2px;background:var(--border2)}
.hdot{width:32px;height:32px;border-radius:10px;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:14px;z-index:1;border:1.5px solid transparent}
.d-cre{background:var(--blu-f);border-color:#BFDBFE}.d-apr{background:var(--grn-f);border-color:#A7F3D0}.d-rej{background:var(--red-f);border-color:#FECACA}.d-val{background:var(--gold-f);border-color:var(--gold-200)}
.hc-lv{font-size:10.5px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--muted)}
.hc-ac{font-size:13.5px;font-weight:700;color:var(--ink)}
.hc-mt{font-size:12px;color:var(--muted);font-weight:500}
.alert{padding:12px 14px;border-radius:10px;font-size:13px;margin-bottom:14px;display:flex;align-items:flex-start;gap:10px;font-weight:500}
.a-success{background:var(--grn-f);border:1.5px solid #A7F3D0;color:#065F46}
.a-warning{background:var(--gold-f);border:1.5px solid var(--gold-200);color:var(--gold-dk)}
.a-danger{background:var(--red-f);border:1.5px solid #FECACA;color:#991B1B}
.a-info{background:var(--blu-f);border:1.5px solid #BFDBFE;color:#1E40AF}
.upzone{border:2px dashed var(--gold-200);border-radius:var(--r);padding:32px;text-align:center;cursor:pointer;transition:all var(--tr);background:var(--gold-f)}
.upzone:hover,.upzone.drag{border-color:var(--gold);background:var(--gold-100)}
.upzone-ico{font-size:40px;margin-bottom:12px}
.upzone-title{font-family:'Sora',sans-serif;font-size:16px;font-weight:700;color:var(--ink)}
.upzone-sub{font-size:12.5px;color:var(--muted);margin-top:4px;font-weight:500}
.dv-head{font-family:'Sora',sans-serif;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--gold-dk);margin-bottom:12px;margin-top:16px;padding-bottom:6px;border-bottom:1.5px solid var(--gold-100)}
.dv-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:10px 20px;margin-bottom:6px}
.dv-lbl{font-size:10.5px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:var(--muted);margin-bottom:2px}
.dv-val{font-size:13.5px;color:var(--ink);font-weight:600}
.dv-val.empty{color:var(--border);font-style:italic;font-size:12px}
.pag{display:flex;align-items:center;gap:8px;padding-top:14px;justify-content:space-between;flex-wrap:wrap}
.pag-info{font-size:12px;color:var(--muted);font-weight:600}
.pag-btns{display:flex;gap:4px}
.pb{width:32px;height:32px;border-radius:8px;border:1.5px solid var(--border);background:var(--white);cursor:pointer;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center;transition:all var(--tr);color:var(--ink3)}
.pb:hover{border-color:var(--gold);color:var(--gold-dk)}.pb.active{background:var(--gold);border-color:var(--gold);color:#fff;box-shadow:0 4px 10px rgba(212,137,10,.3)}
.pb:disabled{opacity:.35;cursor:default}
.cm-sect-nav{width:200px;background:var(--cream);border-right:1.5px solid var(--border2);padding:12px 8px;overflow-y:auto;flex-shrink:0}
.cm-sn-item{padding:9px 12px;cursor:pointer;font-size:12.5px;font-weight:600;border-radius:8px;transition:all .15s;display:flex;align-items:center;gap:8px;color:var(--muted);margin-bottom:2px}
.cm-sn-item:hover{background:var(--gold-f);color:var(--gold-dk)}.cm-sn-item.active{background:var(--white);color:var(--gold-dk);font-weight:700;box-shadow:var(--sh1);border:1px solid var(--gold-200)}
.login-wrap{min-height:100vh;display:flex;background:var(--cream)}
.login-left{flex:1;background:linear-gradient(160deg,#1A1200 0%,#2E2000 40%,#4A3500 70%,#6B5200 100%);display:flex;flex-direction:column;justify-content:center;align-items:center;padding:48px;position:relative;overflow:hidden}
.login-left::before{content:'';position:absolute;inset:0;background:radial-gradient(circle at 30% 50%,rgba(212,137,10,.15) 0%,transparent 60%)}
.login-brand-logo{width:80px;height:80px;background:linear-gradient(135deg,var(--gold-lt),var(--gold));border-radius:24px;display:flex;align-items:center;justify-content:center;font-family:'Sora',sans-serif;font-weight:800;color:#fff;font-size:32px;margin-bottom:24px;box-shadow:0 16px 40px rgba(212,137,10,.4);position:relative;z-index:1}
.login-brand-title{font-family:'Sora',sans-serif;font-size:28px;font-weight:800;color:#FFE599;margin-bottom:8px;letter-spacing:-.5px;position:relative;z-index:1;text-align:center}
.login-brand-sub{font-size:14px;color:rgba(255,220,100,.6);text-align:center;font-weight:500;position:relative;z-index:1;line-height:1.6;max-width:280px}
.login-features{margin-top:40px;display:flex;flex-direction:column;gap:14px;position:relative;z-index:1;width:100%;max-width:300px}
.lf-item{display:flex;align-items:center;gap:12px;font-size:13px;color:rgba(255,220,100,.75);font-weight:500}
.lf-dot{width:8px;height:8px;border-radius:50%;background:var(--gold-lt);flex-shrink:0}
.login-right{width:460px;display:flex;flex-direction:column;justify-content:center;padding:48px;background:var(--white)}
.login-form-title{font-family:'Sora',sans-serif;font-size:24px;font-weight:800;color:var(--ink);margin-bottom:6px;letter-spacing:-.5px}
.login-form-sub{font-size:14px;color:var(--muted);margin-bottom:32px;font-weight:500}
.login-label{font-size:11.5px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--ink3);margin-bottom:6px;display:block}
.login-field{margin-bottom:18px}
.toast-wrap{position:fixed;top:76px;right:18px;z-index:9999;display:flex;flex-direction:column;gap:8px;pointer-events:none}
.toast{padding:12px 16px;border-radius:12px;font-size:13px;display:flex;align-items:flex-start;gap:10px;box-shadow:var(--sh3);animation:sUp .25s ease;min-width:280px;max-width:420px;pointer-events:all;font-weight:600;border:1.5px solid transparent}
.t-success{background:var(--white);border-color:#A7F3D0;color:#065F46}
.t-danger{background:var(--white);border-color:#FECACA;color:#991B1B}
.t-warning{background:var(--white);border-color:var(--gold-200);color:var(--gold-dk)}
.pass-wrap{position:relative}
.pass-toggle{position:absolute;right:12px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;color:var(--muted);font-size:15px;padding:2px;transition:color var(--tr)}
.pass-toggle:hover{color:var(--gold-dk)}
.b-blocked-order{background:#FFF0E0;color:#C2410C;border:1.5px solid #FDBA74}.b-completed{background:#F0F9FF;color:#0369A1;border:1.5px solid #BAE6FD}.b-sendback{background:#FEF3C7;color:#92400E;border:1.5px solid #FCD34D}
.csv-ta{width:100%;height:360px;font-family:monospace;font-size:12px;border:1.5px solid var(--border);border-radius:10px;padding:12px;resize:vertical;background:var(--cream);color:var(--ink);line-height:1.65}
.csv-ta:focus{outline:none;border-color:var(--gold);box-shadow:0 0 0 3px rgba(212,137,10,.1)}
.prog-bar{height:6px;background:var(--border2);border-radius:6px;overflow:hidden;margin-top:6px}
.prog-fill{height:100%;border-radius:6px;transition:width .5s ease}
.att-zone{border:2px dashed var(--border);border-radius:12px;padding:28px 20px;text-align:center;cursor:pointer;transition:all .2s;background:var(--cream);position:relative}
.att-zone:hover,.att-zone.drag{border-color:var(--gold);background:var(--gold-f)}
.att-zone-ico{font-size:32px;margin-bottom:8px}
.att-zone-title{font-family:'Sora',sans-serif;font-size:14px;font-weight:700;color:var(--ink);margin-bottom:4px}
.att-zone-sub{font-size:12px;color:var(--muted);font-weight:500}
.att-list{display:flex;flex-direction:column;gap:8px;margin-top:14px}
.att-item{display:flex;align-items:center;gap:12px;padding:10px 14px;background:#fff;border:1.5px solid var(--border2);border-radius:10px;transition:all .15s}
.att-item:hover{border-color:var(--gold-200);box-shadow:var(--sh1)}
.att-icon{width:36px;height:36px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0}
.att-name{font-size:13px;font-weight:700;color:var(--ink);flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.att-meta{font-size:11px;color:var(--muted);font-weight:500;margin-top:2px}
.att-actions{display:flex;gap:6px;flex-shrink:0}
.att-empty{text-align:center;padding:32px 20px;color:var(--muted)}
.att-empty-ico{font-size:40px;margin-bottom:10px}
.att-count-badge{background:var(--gold);color:#fff;font-size:10px;font-weight:700;border-radius:12px;padding:2px 7px;margin-left:6px;vertical-align:middle}
`;

// ═══ HOOKS — defined before any component that uses them ════
function useToast() {
  const [t, setT] = useState(null);
  const show = (msg, type = "success") => { setT({ msg, type }); setTimeout(() => setT(null), 3200); };
  return [t, show];
}

function useExport() {
  const [modal, setModal] = useState(null);
  const [toast, showToast] = useToast();

  const tryDownload = (csvContent, filename) => {
    // Strategy 1: blob + anchor
    try {
      const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = filename; a.style.display = "none";
      document.body.appendChild(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(url); document.body.removeChild(a); }, 1000);
      return "download";
    } catch {}
    // Strategy 2: data URI new tab
    try {
      const uri = "data:text/csv;charset=utf-8," + encodeURIComponent(csvContent);
      const w = window.open(uri, "_blank");
      if (w) return "opened";
    } catch {}
    // Strategy 3: clipboard API
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(csvContent);
        return "clipboard";
      }
    } catch {}
    return "modal";
  };

  const exportCsv = (rows, cols, filename, title) => {
    const csv = buildCsv(rows, cols);
    const result = tryDownload(csv, filename);
    if (result === "download") showToast(`"${filename}" downloading…`, "success");
    else if (result === "opened") showToast("File opened in new tab — save with Ctrl+S", "success");
    else if (result === "clipboard") showToast("CSV copied to clipboard — paste into Excel", "success");
    else setModal({ title: title || filename, csv });
  };

  const exportTemplate = (cols, filename) => {
    const csv = buildTemplateCsv(cols);
    const result = tryDownload(csv, filename);
    if (result === "download") showToast(`Template downloading…`, "success");
    else if (result === "clipboard") showToast("Template copied to clipboard", "success");
    else setModal({ title: `Template — ${filename}`, csv });
  };

  const ExportModal = modal
    ? <CsvExportModal title={modal.title} csvContent={modal.csv} onClose={() => setModal(null)} />
    : null;

  const ToastEl = toast ? <Toast msg={toast.msg} type={toast.type} /> : null;

  return { exportCsv, exportTemplate, ExportModal, ToastEl };
}

// ═══ SHARED UI COMPONENTS ════════════════════════════════════
function Toast({ msg, type }) {
  return (
    <div className="toast-wrap">
      <div className={`toast ${type === "danger" ? "t-danger" : type === "warning" ? "t-warning" : "t-success"}`}>
        <span style={{ fontSize: 18 }}>{type === "danger" ? "❌" : type === "warning" ? "⚠️" : "✅"}</span>
        <span>{msg}</span>
      </div>
    </div>
  );
}

function CsvExportModal({ title, csvContent, onClose }) {
  const tableRef = useRef();
  const taRef = useRef();
  const [copied, setCopied] = useState(false);
  const [activeTab, setActiveTab] = useState("table");

  const lines = csvContent.trim().split("\n");
  const headers = lines[0].split(",").map(h => h.replace(/^"|"$/g, "").trim());
  const rows = lines.slice(1).map(ln => {
    const cols = []; let cur = ""; let inQ = false;
    for (let i = 0; i < ln.length; i++) {
      if (ln[i] === '"') { inQ = !inQ; }
      else if (ln[i] === "," && !inQ) { cols.push(cur.trim()); cur = ""; }
      else { cur += ln[i]; }
    }
    cols.push(cur.trim()); return cols;
  });

  const copyTable = () => {
    const el = tableRef.current;
    if (!el) return copyCsv();
    try {
      const range = document.createRange();
      range.selectNode(el);
      window.getSelection().removeAllRanges();
      window.getSelection().addRange(range);
      document.execCommand("copy");
      window.getSelection().removeAllRanges();
      setCopied(true); setTimeout(() => setCopied(false), 2500);
    } catch { copyCsv(); }
  };

  const copyCsv = () => {
    const doCopy = () => {
      if (taRef.current) {
        taRef.current.select();
        try { document.execCommand("copy"); setCopied(true); setTimeout(() => setCopied(false), 2500); } catch {}
      }
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(csvContent)
        .then(() => { setCopied(true); setTimeout(() => setCopied(false), 2500); })
        .catch(doCopy);
    } else doCopy();
  };

  const doCopy = () => activeTab === "table" ? copyTable() : copyCsv();

  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 960 }}>
        <div className="m-hdr">
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div className="m-hdr-accent">📊</div>
            <div>
              <div className="m-title">Export — {title}</div>
              <div className="m-sub">{rows.length} rows · {headers.length} columns</div>
            </div>
          </div>
          <button className="m-close" onClick={onClose}>×</button>
        </div>
        <div className="m-body">
          {/* 3-step guide */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 18 }}>
            {[
              { n: "1", icon: "📋", title: "Copy Table", desc: 'Click "Copy for Excel" below' },
              { n: "2", icon: "📂", title: "Open Excel / Sheets", desc: "Open a blank workbook, click cell A1" },
              { n: "3", icon: "✅", title: "Paste & Done", desc: "Press Ctrl+V — all columns fill in automatically" },
            ].map(s => (
              <div key={s.n} style={{ background: "var(--cream)", borderRadius: 10, padding: "12px 14px", border: "1.5px solid var(--border2)", display: "flex", gap: 10 }}>
                <div style={{ width: 28, height: 28, borderRadius: 8, background: "var(--gold)", color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 13, flexShrink: 0 }}>{s.n}</div>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 13, color: "var(--ink)", marginBottom: 2 }}>{s.icon} {s.title}</div>
                  <div style={{ fontSize: 11.5, color: "var(--muted)", fontWeight: 500, lineHeight: 1.4 }}>{s.desc}</div>
                </div>
              </div>
            ))}
          </div>

          {/* Controls */}
          <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap", alignItems: "center" }}>
            <button className="btn btn-primary" onClick={doCopy} style={{ minWidth: 200 }}>
              {copied ? "✅ Copied! Now paste in Excel" : activeTab === "table" ? "📋 Copy for Excel / Sheets" : "📋 Copy CSV Text"}
            </button>
            <div style={{ display: "flex", background: "var(--cream)", borderRadius: 8, border: "1.5px solid var(--border2)", overflow: "hidden" }}>
              {[{ k: "table", l: "📊 Table View" }, { k: "csv", l: "📄 CSV Text" }].map(t => (
                <button key={t.k} onClick={() => setActiveTab(t.k)} style={{ padding: "7px 14px", fontSize: 12, fontWeight: 700, fontFamily: "inherit", border: "none", cursor: "pointer", background: activeTab === t.k ? "var(--gold)" : "transparent", color: activeTab === t.k ? "#fff" : "var(--muted)", transition: "all .15s" }}>{t.l}</button>
              ))}
            </div>
            <span style={{ marginLeft: "auto", fontSize: 12, color: "var(--muted)", fontWeight: 600 }}>
              {rows.length} rows · {headers.length} cols · {(csvContent.length / 1024).toFixed(1)} KB
            </span>
          </div>

          {/* Table view */}
          {activeTab === "table" && (
            <div style={{ maxHeight: 340, overflow: "auto", border: "1.5px solid var(--border)", borderRadius: 10 }}>
              <table ref={tableRef} style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, fontFamily: "'Plus Jakarta Sans',sans-serif" }}>
                <thead>
                  <tr style={{ background: "#FFF0CC", position: "sticky", top: 0 }}>
                    {headers.map((h, i) => (
                      <th key={i} style={{ padding: "8px 10px", textAlign: "left", fontWeight: 700, fontSize: 11, textTransform: "uppercase", letterSpacing: ".4px", color: "#7A5500", borderBottom: "2px solid #FFD166", whiteSpace: "nowrap", borderRight: "1px solid #FFE08A" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.map((row, i) => (
                    <tr key={i} style={{ background: i % 2 === 0 ? "#fff" : "#FFFDF5" }}>
                      {headers.map((_, j) => (
                        <td key={j} style={{ padding: "7px 10px", borderBottom: "1px solid #F0EDE4", borderRight: "1px solid #F5EDD5", fontSize: 12, color: "#2A2010", whiteSpace: "nowrap", maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis" }}>
                          {row[j] ?? ""}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* CSV text view */}
          {activeTab === "csv" && (
            <textarea ref={taRef} readOnly value={csvContent} className="csv-ta" onFocus={e => e.target.select()} style={{ height: 320 }} />
          )}

          <div style={{ marginTop: 12, padding: "10px 14px", background: "var(--gold-f)", borderRadius: 8, border: "1.5px solid var(--gold-200)", fontSize: 12, color: "var(--gold-dk)", fontWeight: 600 }}>
            💡 <strong>Excel tip:</strong> Paste with Ctrl+V into Excel — data fills separate columns automatically. No import wizard needed.
          </div>
        </div>
        <div className="m-footer">
          <button className="btn btn-secondary btn-sm" onClick={onClose}>Close</button>
          <button className="btn btn-primary" onClick={doCopy}>
            {copied ? "✅ Copied! Open Excel & paste" : "📋 Copy for Excel / Sheets"}
          </button>
        </div>
      </div>
    </div>
  );
}

function ConfirmDelete({ label, onConfirm, onClose }) {
  return (
    <div className="overlay" style={{ zIndex: 700 }}>
      <div className="modal modal-xs">
        <div className="m-hdr">
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div className="m-hdr-accent" style={{ background: "linear-gradient(135deg,#EF4444,#DC2626)" }}>🗑</div>
            <div><div className="m-title">Delete Record</div><div className="m-sub">This cannot be undone</div></div>
          </div>
          <button className="m-close" onClick={onClose}>×</button>
        </div>
        <div className="m-body" style={{ textAlign: "center", padding: "28px 24px" }}>
          <div style={{ fontSize: 48, marginBottom: 12 }}>⚠️</div>
          <div style={{ fontFamily: "'Sora',sans-serif", fontSize: 15, fontWeight: 700, marginBottom: 8 }}>Are you sure?</div>
          <div style={{ fontSize: 13, color: "var(--muted)", fontWeight: 500 }}>You are about to delete <strong style={{ color: "var(--ink)" }}>{label}</strong>.</div>
        </div>
        <div className="m-footer">
          <button className="btn btn-secondary btn-sm" onClick={onClose}>Cancel</button>
          <button className="btn btn-danger btn-sm" onClick={() => { onConfirm(); onClose(); }}>Yes, Delete</button>
        </div>
      </div>
    </div>
  );
}

function ImportModal({ title, colDefs, previewCols, onImport, onClose }) {
  const [drag, setDrag] = useState(false);
  const [preview, setPreview] = useState(null);
  const [fname, setFname] = useState("");
  const [err, setErr] = useState("");
  const fileRef = useRef();
  const handleFile = file => {
    if (!file) return;
    if (!file.name.endsWith(".csv")) { setErr("Please upload a CSV file."); return; }
    setErr(""); setFname(file.name);
    const rd = new FileReader();
    rd.onload = e => { const rec = parseCsvToObjects(e.target.result, colDefs); if (!rec.length) { setErr("No valid records found."); return; } setPreview(rec); };
    rd.readAsText(file);
  };
  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 660 }}>
        <div className="m-hdr">
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div className="m-hdr-accent">📤</div>
            <div><div className="m-title">Import — {title}</div><div className="m-sub">Upload CSV with matching headers</div></div>
          </div>
          <button className="m-close" onClick={onClose}>×</button>
        </div>
        <div className="m-body">
          <div className={`upzone${drag ? " drag" : ""}`} onClick={() => fileRef.current.click()}
            onDragOver={e => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)}
            onDrop={e => { e.preventDefault(); setDrag(false); handleFile(e.dataTransfer.files[0]); }}>
            <div className="upzone-ico">📂</div>
            <div className="upzone-title">{fname || "Drop CSV file here"}</div>
            <div className="upzone-sub">{fname ? "File loaded — see preview below" : "or click to browse"}</div>
            <input ref={fileRef} type="file" accept=".csv" style={{ display: "none" }} onChange={e => handleFile(e.target.files[0])} />
          </div>
          {err && <div className="alert a-danger" style={{ marginTop: 12 }}><span>⚠</span>{err}</div>}
          {preview && (
            <div style={{ marginTop: 16 }}>
              <div style={{ background: "var(--grn-f)", color: "#065F46", padding: "10px 14px", borderRadius: 10, display: "inline-flex", alignItems: "center", gap: 10, marginBottom: 12, border: "1.5px solid #A7F3D0", fontWeight: 700, fontSize: 13 }}>
                ✅ {preview.length} records ready
              </div>
              <div style={{ maxHeight: 180, overflow: "auto", border: "1.5px solid var(--border2)", borderRadius: 10 }}>
                <table><thead><tr>{previewCols.map(c => <th key={c}>{c}</th>)}</tr></thead>
                  <tbody>{preview.slice(0, 5).map((r, i) => <tr key={i}>{previewCols.map(c => <td key={c} style={{ fontSize: 12 }}>{r[c] || "—"}</td>)}</tr>)}</tbody></table>
              </div>
            </div>
          )}
        </div>
        <div className="m-footer">
          <button className="btn btn-secondary btn-sm" onClick={onClose}>Cancel</button>
          <button className="btn btn-primary btn-sm" disabled={!preview} onClick={() => { onImport(preview); onClose(); }}>Import {preview ? preview.length : 0} Records</button>
        </div>
      </div>
    </div>
  );
}

function StatusBadge({ status, blocked }) {
  if (blocked) return <span className="badge b-blocked-order">🔒 Blocked</span>;
  const m = { Pending: "b-pending", Approved: "b-approved", Rejected: "b-rejected", Validated: "b-validated", "Final Approved": "b-approved", "Sent Back": "b-sendback", Completed: "b-completed" };
  const ic = { Pending: "🕐", Approved: "✅", Rejected: "❌", Validated: "🔍", "Sent Back": "↩", Completed: "🏁" };
  return <span className={`badge ${m[status] || "b-pending"}`}>{ic[status] || ""} {status}</span>;
}

function HistLine({ history, users }) {
  const dc = a => ({ Created: "d-cre", Approved: "d-apr", Validated: "d-apr", Rejected: "d-rej", Resubmitted: "d-val" }[a] || "d-val");
  const ic = a => ({ Created: "📝", Approved: "✅", Rejected: "❌", Validated: "🔍", Modified: "✏️", Escalated: "📤", "Sent Back": "↩", Resubmitted: "🚀", Blocked: "🔒", Unblocked: "🔓" }[a] || "•");
  return (
    <div className="htl">{history.map((h, i) => (
      <div className="hi" key={i}>
        <div className={`hdot ${dc(h.action)}`}>{ic(h.action)}</div>
        <div style={{ flex: 1 }}>
          <div className="hc-lv">{h.level}</div>
          <div className="hc-ac">{h.action} · <span style={{ color: "var(--gold-dk)" }}>{users.find(u => u.id === h.by)?.name || h.by}</span></div>
          {h.remark && <div className="hc-mt" style={{ background: "var(--cream)", padding: "4px 10px", borderRadius: 6, marginTop: 4, fontSize: 12 }}>"{h.remark}"</div>}
          <div className="hc-mt">{h.time}</div>
        </div>
      </div>
    ))}</div>
  );
}

// ═══ LOGIN ═════════════════════════════════════════════════
function LoginPage({ users, onLogin }) {
  const [uid, setUid] = useState(users[0]?.id || "");
  const [pw, setPw] = useState(""); const [showPw, setShowPw] = useState(false); const [err, setErr] = useState(""); const [loading, setLoading] = useState(false);
  const handleLogin = () => {
    const u = users.find(u => u.id === uid);
    if (!u) { setErr("User not found."); return; }
    if (u.status === "Inactive") { setErr("Account inactive. Contact Admin."); return; }
    if (u.password !== pw) { setErr("Incorrect password."); return; }
    setLoading(true); setTimeout(() => { setLoading(false); onLogin(u); }, 600);
  };
  const BU_LOGO = "data:image/png;base64,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";
  return (
    <div style={{ minHeight:"100vh", display:"flex", fontFamily:"'Plus Jakarta Sans',sans-serif", background:"#fff" }}>
      <style>{CSS}</style>
      {/* Left panel — red brand side */}
      <div style={{ flex:"0 0 48%", background:"linear-gradient(160deg,#C0142E 0%,#8B0E20 100%)", display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", padding:"48px 56px", position:"relative", overflow:"hidden" }}>
        {/* Decorative circles */}
        <div style={{ position:"absolute", top:-80, right:-80, width:300, height:300, borderRadius:"50%", background:"rgba(255,255,255,.06)" }}/>
        <div style={{ position:"absolute", bottom:-60, left:-60, width:240, height:240, borderRadius:"50%", background:"rgba(255,255,255,.05)" }}/>
        {/* Logo */}
        {/* Logo card - clips black bottom bar via clipPath */}
        <div style={{ background:"#fff", borderRadius:20, overflow:"hidden", marginBottom:28, boxShadow:"0 20px 60px rgba(0,0,0,.3)", maxWidth:360, width:"100%" }}>
          <img src={BU_LOGO} alt="Birla Uttam Cement" style={{ width:"100%", display:"block", clipPath:"inset(0 0 27.5% 0)", marginBottom:"-0.5px" }}/>
        </div>
        {/* Company name in white below logo */}
        <div style={{ textAlign:"center", marginBottom:36 }}>
          <div style={{ fontSize:26, fontWeight:900, color:"#fff", letterSpacing:.5, lineHeight:1.2, fontFamily:"'Sora',sans-serif", marginBottom:6 }}>
            Mangalam Cement Limited
          </div>
          <div style={{ fontSize:13.5, color:"rgba(255,255,255,.75)", fontWeight:500, letterSpacing:.3 }}>
            Non-Trade Price Approval System
          </div>
          <div style={{ marginTop:6, fontSize:12, color:"rgba(255,255,255,.55)", fontStyle:"italic" }}>
            Streamlining field sales across Rajasthan & Uttar Pradesh
          </div>
        </div>
        {/* Feature list */}
        <div style={{ width:"100%", maxWidth:340 }}>
          {["Multi-level approval workflow","Real-time NCR monitoring","Customer master management","Comprehensive analytics & reports","Role-based access control"].map(f => (
            <div key={f} style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10 }}>
              <div style={{ width:7, height:7, borderRadius:"50%", background:"rgba(255,255,255,.6)", flexShrink:0 }}/>
              <span style={{ fontSize:13, color:"rgba(255,255,255,.82)", fontWeight:500 }}>{f}</span>
            </div>
          ))}
        </div>
        {/* Bottom tag */}
        <div style={{ position:"absolute", bottom:20, fontSize:11, color:"rgba(255,255,255,.35)", letterSpacing:.4 }}>
          © 2024 Mangalam Cement Limited. All rights reserved.
        </div>
      </div>

      {/* Right panel — white login form */}
      <div style={{ flex:1, display:"flex", alignItems:"center", justifyContent:"center", background:"#fff", padding:"40px 32px" }}>
        <div style={{ width:"100%", maxWidth:420 }}>
          {/* Top logo small */}
          <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:36 }}>
            <div style={{ background:"linear-gradient(135deg,#C0142E,#8B0E20)", borderRadius:12, width:44, height:44, display:"flex", alignItems:"center", justifyContent:"center", boxShadow:"0 6px 20px rgba(192,20,46,.3)" }}>
              <span style={{ fontSize:20 }}>🏭</span>
            </div>
            <div>
              <div style={{ fontSize:15, fontWeight:800, color:"#1A0000", letterSpacing:.2 }}>Mangalam Cement Limited</div>
              <div style={{ fontSize:11, color:"#888", fontWeight:500 }}>NTPAS — Price Approval System</div>
            </div>
          </div>

          <div style={{ marginBottom:28 }}>
            <div style={{ fontSize:28, fontWeight:900, color:"#1A0000", marginBottom:6, fontFamily:"'Sora',sans-serif" }}>Welcome back 👋</div>
            <div style={{ fontSize:14, color:"#888" }}>Sign in to access your account</div>
          </div>

          {err && (
            <div style={{ display:"flex", alignItems:"center", gap:8, background:"#FEF2F2", border:"1.5px solid #FECACA", borderRadius:10, padding:"10px 14px", marginBottom:20, fontSize:13, color:"#DC2626", fontWeight:600 }}>
              <span>⚠️</span><span>{err}</span>
            </div>
          )}

          <div style={{ marginBottom:18 }}>
            <label style={{ display:"block", fontSize:12.5, fontWeight:700, color:"#444", marginBottom:7, letterSpacing:.3, textTransform:"uppercase" }}>User ID / Name</label>
            <select value={uid} onChange={e => { setUid(e.target.value); setErr(""); setPw(""); }}
              style={{ width:"100%", padding:"12px 14px", border:"1.5px solid #E5E7EB", borderRadius:10, fontSize:14, color:"#1A0000", fontFamily:"inherit", outline:"none", background:"#FAFAFA", cursor:"pointer" }}>
              {users.filter(u => u.status !== "Inactive").map(u => <option key={u.id} value={u.id}>{u.name} — {u.role}</option>)}
            </select>
          </div>

          <div style={{ marginBottom:24 }}>
            <label style={{ display:"block", fontSize:12.5, fontWeight:700, color:"#444", marginBottom:7, letterSpacing:.3, textTransform:"uppercase" }}>Password</label>
            <div style={{ position:"relative" }}>
              <input type={showPw?"text":"password"} value={pw}
                onChange={e => { setPw(e.target.value); setErr(""); }}
                onKeyDown={e => e.key==="Enter" && handleLogin()}
                placeholder="Enter your password"
                style={{ width:"100%", padding:"12px 44px 12px 14px", border:"1.5px solid #E5E7EB", borderRadius:10, fontSize:14, color:"#1A0000", fontFamily:"inherit", outline:"none", background:"#FAFAFA", boxSizing:"border-box" }}/>
              <button onClick={()=>setShowPw(p=>!p)}
                style={{ position:"absolute", right:12, top:"50%", transform:"translateY(-50%)", background:"none", border:"none", cursor:"pointer", fontSize:16, padding:0 }}>
                {showPw?"🙈":"👁"}
              </button>
            </div>
            <div style={{ fontSize:11, color:"#AAA", marginTop:5, fontStyle:"italic" }}>Default: firstname + numbers (e.g. amit123)</div>
          </div>

          <button onClick={handleLogin} disabled={loading}
            style={{ width:"100%", padding:"13px", background:loading?"#ccc":"linear-gradient(135deg,#C0142E,#8B0E20)", color:"#fff", border:"none", borderRadius:11, fontSize:15, fontWeight:800, cursor:loading?"not-allowed":"pointer", fontFamily:"inherit", letterSpacing:.3, boxShadow:"0 6px 20px rgba(192,20,46,.3)", transition:"all .2s" }}>
            {loading?"Signing in…":"Sign In →"}
          </button>

          {/* Demo credentials */}
          <div style={{ marginTop:28, background:"#FFF5F5", border:"1.5px solid #FECACA", borderRadius:12, padding:"14px 16px" }}>
            <div style={{ fontSize:11, fontWeight:800, color:"#C0142E", marginBottom:10, letterSpacing:.5, textTransform:"uppercase" }}>Quick Demo Access</div>
            {[{ id:"ADM001", pw:"admin123", label:"Admin" }, { id:"SO001", pw:"amit123", label:"Sales Officer" }, { id:"ZH001", pw:"narendra123", label:"Zonal Head" }].map(d => (
              <div key={d.id} style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:7 }}>
                <span style={{ fontSize:12.5, color:"#555", fontWeight:600 }}>{d.label}</span>
                <button onClick={() => { setUid(d.id); setPw(d.pw); setErr(""); }}
                  style={{ background:"#C0142E", color:"#fff", border:"none", borderRadius:7, padding:"4px 12px", fontSize:11, cursor:"pointer", fontWeight:700, fontFamily:"inherit" }}>
                  Use
                </button>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

// ═══ DASHBOARD ════════════════════════════════════════════
function Dashboard({ requests, customers, users, user, setPage, setActiveTab, salesEntries }) {
  const cUser=(users||[]).find(u=>u.id===user.id)||user;
  const ur=requests.filter(r=>user.role==="Admin"||user.role==="Sales & Accounts"||user.role==="Zonal Head"||(user.assignedCustomers||[]).includes(r.customerCode));

  // ── Shared mini-components for analytics ─────────────────
  function AKpi({icon,val,label,trend,trendTxt,color}) {
    return (
      <div style={{background:"var(--white)",border:"1.5px solid var(--border)",borderRadius:10,padding:"12px 14px",position:"relative",overflow:"hidden"}}>
        <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:color||"var(--gold)"}}/>
        <div style={{fontSize:20,marginBottom:6}}>{icon}</div>
        <div style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:String(val).length>7?16:22,color:"var(--ink)"}}>{val}</div>
        <div style={{fontSize:10.5,fontWeight:700,color:"var(--muted)",textTransform:"uppercase",letterSpacing:.4,marginBottom:3}}>{label}</div>
        {trendTxt&&<div style={{fontSize:10.5,fontWeight:700,color:trend==="up"?"var(--grn)":trend==="down"?"var(--red)":"var(--muted)"}}>{trend==="up"?"↑":trend==="down"?"↓":"→"} {trendTxt}</div>}
      </div>
    );
  }
  function ABar({items,maxVal,colorFn}) {
    const mx=maxVal||Math.max(...items.map(i=>i.val),1);
    return (
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {items.map((it,i)=>(
          <div key={i} style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:90,fontSize:11.5,fontWeight:600,color:"var(--ink)",flexShrink:0,textAlign:"right"}}>{it.label}</div>
            <div style={{flex:1,background:"var(--border)",borderRadius:4,height:22,overflow:"hidden"}}>
              <div style={{height:"100%",width:(aPct(it.val,mx))+"%",background:colorFn?colorFn(it,i):"var(--gold)",borderRadius:4,display:"flex",alignItems:"center",paddingLeft:6}}>
                {it.val>0&&<span style={{fontSize:11,fontWeight:700,color:"#fff"}}>{it.val}</span>}
              </div>
            </div>
            <div style={{width:50,fontSize:11.5,fontWeight:700,color:"var(--muted)",textAlign:"right",flexShrink:0}}>{it.sub||it.val}</div>
          </div>
        ))}
      </div>
    );
  }
  function AGrid({cols,children}) {
    return <div style={{display:"grid",gridTemplateColumns:"repeat("+cols+",1fr)",gap:12,marginBottom:14}}>{children}</div>;
  }
  function ACard({title,sub,children,style}) {
    return (
      <div className="card" style={style}>
        {(title||sub)&&<div className="card-hdr"><div><div className="card-title">{title}</div>{sub&&<div className="card-sub">{sub}</div>}</div></div>}
        <div className="card-body">{children}</div>
      </div>
    );
  }
  function APh({title,sub,right}) {
    return <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:18}}><div><div style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:18,color:"var(--ink)"}}>{title}</div><div style={{fontSize:12,color:"var(--muted)",marginTop:3}}>{sub}</div></div>{right&&<div style={{textAlign:"right"}}>{right}</div>}</div>;
  }

  // ── Analytics Dashboards ──────────────────────────────────
  function AnExec() {
    const app=ur.filter(r=>r.status==="Approved"),pnd=ur.filter(r=>r.status==="Pending"),rej=ur.filter(r=>r.status==="Rejected"),blk=ur.filter(r=>r.blocked);
    const vol=aSum(app,"qty"),avgNCR=aAvg(ur,"ncrPmt"),negNCR=ur.filter(r=>r.ncrPmt<0).length;
    const rev=(aSum(app,"orderPrice")*aSum(app,"qty")/1000000).toFixed(1);
    const gradeData=GRADES_LIST.map(g=>({label:g,val:aSum(app.filter(r=>r.grade===g),"qty"),sub:app.filter(r=>r.grade===g).length+" req"}));
    const regions=[...new Set(ur.map(r=>r.region).filter(Boolean))];
    const regionData=regions.map(r=>({label:r.length>14?r.slice(0,12)+"…":r,val:ur.filter(x=>x.region===r).length})).sort((a,b)=>b.val-a.val).slice(0,5);
    return (
      <div>
        <APh title="Executive Overview" sub={"All-time KPIs · "+ur.length+" requests"} right={<div><div style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:24,color:aPct(app.length,ur.length)>60?"var(--grn)":"var(--ora)"}}>{aPct(app.length,ur.length)}%</div><div style={{fontSize:11,color:"var(--muted)",fontWeight:600}}>Approval Rate</div></div>}/>
        <AGrid cols={5}><AKpi icon="📋" val={ur.length}           label="Total Requests"    trend="neu" trendTxt="All time"              color="var(--gold)"/><AKpi icon="✅" val={app.length}  label="Approved"         trend="up"  trendTxt={aPct(app.length,ur.length)+"%"} color="var(--grn)"/><AKpi icon="🕐" val={pnd.length}  label="Pending"          trend={pnd.length>3?"down":"up"} trendTxt="In pipeline" color="#F59E0B"/><AKpi icon="❌" val={rej.length}  label="Rejected"         trend="neu" trendTxt={aPct(rej.length,ur.length)+"%"} color="var(--red)"/><AKpi icon="🔒" val={blk.length}  label="Blocked"          trend="neu" trendTxt="Admin blocked"         color="var(--ora)"/></AGrid>
        <AGrid cols={4}><AKpi icon="📊" val={"₹"+avgNCR}          label="Avg NCR/MT"       trend={avgNCR<0?"down":"up"} trendTxt="Per tonne" color={avgNCR<0?"var(--red)":"var(--grn)"}/><AKpi icon="⚠️" val={negNCR}            label="Negative NCR"     trend="down" trendTxt={aPct(negNCR,ur.length)+"%"}  color="var(--red)"/><AKpi icon="📦" val={(vol/1000).toFixed(1)+"K"} label="Approved Vol MT"  trend="up"  trendTxt="Metric tonnes"            color="#7C3AED"/><AKpi icon="💰" val={"₹"+rev+"Cr"}     label="Revenue Impact"   trend="up"  trendTxt="Approved orders"          color="#2563EB"/></AGrid>
        <AGrid cols={2}>
          <ACard title="Grade-wise Volume (MT)" sub="Approved orders"><ABar items={gradeData} colorFn={(_,i)=>["var(--gold)","#2563EB","#7C3AED"][i]||"var(--gold)"}/></ACard>
          <ACard title="Region-wise Requests" sub="All statuses"><ABar items={regionData} colorFn={(_,i)=>["var(--gold)","#059669","#2563EB","#7C3AED","#EA580C"][i]||"var(--gold)"}/></ACard>
        </AGrid>
      </div>
    );
  }

  function AnWorkflow() {
    const LEVELS=["Sales Officer","Area Sales Manager","Regional Head","Zonal Head","Sales & Accounts","Admin"];
    const SHORT=["SO","ASM","RH","ZH","S&A","ADM"];
    const COLS=["#F59E0B","#3B82F6","#8B5CF6","#059669","#0891B2","var(--gold)"];
    const levelData=LEVELS.map((lvl,i)=>({name:lvl,short:SHORT[i],color:COLS[i],pending:ur.filter(r=>r.status==="Pending"&&r.currentLevel===lvl).length,total:ur.filter(r=>r.currentLevel===lvl).length}));
    const maxPend=Math.max(...levelData.map(l=>l.pending),1);
    return (
      <div>
        <APh title="Approval Workflow Dashboard" sub="Funnel · Stage breakdown · Level metrics"/>
        <AGrid cols={2}>
          <ACard title="Pending Count by Level" sub="Approval funnel">
            {levelData.map(l=>(
              <div key={l.name} style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
                <div style={{width:32,fontSize:11.5,fontWeight:800,color:l.color,textAlign:"right",flexShrink:0}}>{l.short}</div>
                <div style={{flex:1,background:"var(--border)",borderRadius:5,height:32,overflow:"hidden"}}>
                  <div style={{height:"100%",width:Math.max(15,aPct(l.pending,maxPend))+"%",background:l.color,borderRadius:5,display:"flex",alignItems:"center",paddingLeft:8}}>
                    <span style={{fontSize:12,fontWeight:700,color:"#fff"}}>{l.pending} pending</span>
                  </div>
                </div>
              </div>
            ))}
          </ACard>
          <ACard title="Level Performance Metrics">
            <div className="tw"><table>
              <thead><tr><th>Level</th><th>Pending</th><th>Total</th><th>Status</th></tr></thead>
              <tbody>{levelData.map(l=>(
                <tr key={l.name}>
                  <td><span style={{fontWeight:700,color:l.color}}>{l.short}</span> {l.name}</td>
                  <td style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:l.pending>2?"var(--red)":"var(--grn)"}}>{l.pending}</td>
                  <td>{l.total}</td>
                  <td><span className={l.pending>2?"badge b-pending":"badge b-approved"}>{l.pending>2?"⚠ Busy":"✅ OK"}</span></td>
                </tr>
              ))}</tbody>
            </table></div>
          </ACard>
        </AGrid>
      </div>
    );
  }

  function AnNCR() {
    const [fGrade,setFGrade]=useState(""); const [fPlant,setFPlant]=useState("");
    const fil=ur.filter(r=>(!fGrade||r.grade===fGrade)&&(!fPlant||r.supplyFrom===fPlant));
    const pos=fil.filter(r=>r.ncrPmt>=0),neg=fil.filter(r=>r.ncrPmt<0),avgN=aAvg(fil,"ncrPmt");
    const worst=[...fil].sort((a,b)=>a.ncrPmt-b.ncrPmt).slice(0,5);
    const best=[...fil].sort((a,b)=>b.ncrPmt-a.ncrPmt).slice(0,5);
    const gradeNCR=GRADES_LIST.map(g=>({label:g,val:aAvg(fil.filter(r=>r.grade===g),"ncrPmt"),sub:"₹"+aAvg(fil.filter(r=>r.grade===g),"ncrPmt")}));
    const plantNCR=PLANTS_LIST.map(p=>({label:p.replace(" Plant",""),val:aAvg(fil.filter(r=>r.supplyFrom===p),"ncrPmt"),sub:"₹"+aAvg(fil.filter(r=>r.supplyFrom===p),"ncrPmt")}));
    return (
      <div>
        <APh title="NCR Analytics Dashboard" sub={"Margin intelligence · "+fil.length+" records"} right={<div><div style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:24,color:avgN<0?"var(--red)":"var(--grn)"}}>₹{avgN}</div><div style={{fontSize:11,color:"var(--muted)",fontWeight:600}}>Avg NCR/MT</div></div>}/>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:14,padding:"10px 14px",background:"var(--cream)",borderRadius:10,border:"1.5px solid var(--border2)"}}>
          <span style={{fontSize:12,fontWeight:700,color:"var(--muted)"}}>Filter:</span>
          <select style={{border:"1.5px solid var(--border2)",borderRadius:7,padding:"4px 8px",fontSize:12,fontFamily:"inherit"}} value={fGrade} onChange={function(e){setFGrade(e.target.value)}}><option value="">All Grades</option>{GRADES_LIST.map(g=><option key={g}>{g}</option>)}</select>
          <select style={{border:"1.5px solid var(--border2)",borderRadius:7,padding:"4px 8px",fontSize:12,fontFamily:"inherit"}} value={fPlant} onChange={function(e){setFPlant(e.target.value)}}><option value="">All Plants</option>{PLANTS_LIST.map(p=><option key={p}>{p}</option>)}</select>
          {(fGrade||fPlant)&&<button onClick={function(){setFGrade("");setFPlant("");}} style={{fontSize:11,border:"none",background:"none",color:"var(--red)",cursor:"pointer",fontWeight:700}}>✕ Clear</button>}
        </div>
        <AGrid cols={4}><AKpi icon="📊" val={"₹"+avgN} label="Avg NCR/MT" trend={avgN<0?"down":"up"} trendTxt="Overall" color={avgN<0?"var(--red)":"var(--grn)"}/><AKpi icon="✅" val={pos.length} label="Positive NCR" trend="up" trendTxt={aPct(pos.length,fil.length)+"%"} color="var(--grn)"/><AKpi icon="❌" val={neg.length} label="Negative NCR" trend="down" trendTxt={aPct(neg.length,fil.length)+"%"} color="var(--red)"/><AKpi icon="⚠️" val={fil.filter(r=>r.ncrPmt<-300).length} label="Critical lt-300" trend="down" trendTxt="Urgent review" color="var(--ora)"/></AGrid>
        <AGrid cols={2}>
          <ACard title="Grade-wise Avg NCR/MT"><ABar items={gradeNCR} colorFn={function(it){return Number(it.val)>=0?"var(--grn)":"var(--red)";}} maxVal={Math.max(...gradeNCR.map(g=>Math.abs(g.val)),1)}/></ACard>
          <ACard title="Plant-wise Avg NCR/MT"><ABar items={plantNCR} colorFn={function(it){return Number(it.val)>=0?"var(--grn)":"var(--red)";}} maxVal={Math.max(...plantNCR.map(p=>Math.abs(p.val)),1)}/></ACard>
        </AGrid>
        <AGrid cols={2}>
          <ACard title="🔴 Worst NCR Deals" sub="Lowest margin requests">
            <div className="tw"><table><thead><tr><th>ID</th><th>Customer</th><th>Grade</th><th>NCR/MT</th></tr></thead>
            <tbody>{worst.map(r=><tr key={r.id}><td style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:"var(--gold-dk)",fontSize:12}}>{r.id}</td><td style={{fontWeight:600,fontSize:12}}>{r.customerName}</td><td><span className="badge b-grade">{r.grade}</span></td><td style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:"var(--red)"}}>₹{Number(r.ncrPmt).toFixed(0)}</td></tr>)}</tbody>
            </table></div>
          </ACard>
          <ACard title="🟢 Best Margin Deals" sub="Top NCR requests">
            <div className="tw"><table><thead><tr><th>ID</th><th>Customer</th><th>Grade</th><th>NCR/MT</th></tr></thead>
            <tbody>{best.map(r=><tr key={r.id}><td style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:"var(--gold-dk)",fontSize:12}}>{r.id}</td><td style={{fontWeight:600,fontSize:12}}>{r.customerName}</td><td><span className="badge b-grade">{r.grade}</span></td><td style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:"var(--grn)"}}>₹{Number(r.ncrPmt).toFixed(0)}</td></tr>)}</tbody>
            </table></div>
          </ACard>
        </AGrid>
      </div>
    );
  }

  function AnPerf() {
    const soList=(users||[]).filter(u=>u.role==="Sales Officer"||u.role==="Area Sales Manager");
    const perfData=soList.map(u=>{const mine=ur.filter(r=>r.createdBy===u.id);const app=mine.filter(r=>r.status==="Approved");return{...u,created:mine.length,approved:app.length,appPct:aPct(app.length,mine.length||1),avgNCR:aAvg(mine,"ncrPmt"),vol:aSum(app,"qty")};}).filter(u=>u.created>0).sort((a,b)=>b.appPct-a.appPct);
    return (
      <div>
        <APh title="Team Performance Dashboard" sub="Individual approval rates · NCR managed · Volume"/>
        <AGrid cols={4}><AKpi icon="👥" val={(users||[]).filter(u=>u.role==="Sales Officer").length} label="Sales Officers" trend="neu" trendTxt="Active" color="var(--gold)"/><AKpi icon="🧑‍💼" val={(users||[]).filter(u=>u.role==="Area Sales Manager").length} label="Area Managers" trend="neu" trendTxt="Active" color="#2563EB"/><AKpi icon="📈" val={perfData.length?Math.max(...perfData.map(p=>p.appPct)):0} label="Best Approval %" trend="up" trendTxt="Top performer" color="var(--grn)"/><AKpi icon="📊" val={perfData.length?"₹"+aAvg(perfData,"avgNCR"):"-"} label="Avg NCR Managed" trend="neu" trendTxt="Across team" color="#7C3AED"/></AGrid>
        <ACard title="Team Leaderboard" sub="Ranked by approval rate">
          {perfData.length===0&&<div style={{padding:24,textAlign:"center",color:"var(--muted)"}}>No performance data yet.</div>}
          {perfData.map(function(u,i){return(
            <div key={u.id} style={{display:"flex",alignItems:"center",gap:12,padding:"10px 0",borderBottom:"1px solid var(--border2)"}}>
              <div style={{width:22,height:22,borderRadius:"50%",background:["var(--gold)","var(--muted)","#CD7F32","var(--border2)"][i]||"var(--border2)",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:11,flexShrink:0}}>{i+1}</div>
              <div style={{width:140}}><div style={{fontWeight:700,fontSize:13}}>{u.name}</div><div style={{fontSize:10.5,color:"var(--muted)"}}>{u.role} · {u.created} requests</div></div>
              <div style={{flex:1}}><div style={{background:"var(--border)",borderRadius:4,height:7,overflow:"hidden"}}><div style={{height:"100%",width:u.appPct+"%",background:"linear-gradient(90deg,var(--gold),var(--gold-lt))",borderRadius:4}}/></div></div>
              <div style={{width:50,textAlign:"right",fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:15,color:"var(--grn)"}}>{u.appPct}%</div>
              <div style={{width:65,textAlign:"right",fontSize:12,fontWeight:700,color:u.avgNCR<0?"var(--red)":"var(--grn)"}}>₹{u.avgNCR}</div>
            </div>
          );})}
        </ACard>
      </div>
    );
  }

  function AnSales() {
    const approved=ur.filter(r=>r.status==="Approved");
    const enriched=approved.map(function(r){const ent=(salesEntries||[]).filter(e=>e.requestId===r.id);const sold=ent.reduce((s,e)=>s+Number(e.invoiceQty||0),0);return{...r,sold,pending:Math.max(0,r.qty-sold),util:aPct(sold,r.qty||1),invCount:ent.length};});
    const totalAppr=aSum(approved,"qty"),totalSold=enriched.reduce((s,r)=>s+r.sold,0);
    const overallUtil=aPct(totalSold,totalAppr);
    const lowUtil=enriched.filter(r=>r.util<50&&r.qty>0).slice(0,6);
    const gradeItems=GRADES_LIST.map(g=>{const ga=enriched.filter(r=>r.grade===g);const ap=aSum(ga,"qty");const so=ga.reduce((s,r)=>s+r.sold,0);return{label:g,val:aPct(so,ap),sub:aPct(so,ap)+"%"};});
    return (
      <div>
        <APh title="Sales vs Approval Dashboard" sub="Approved qty utilization · Gap analysis" right={<div><div style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:24,color:overallUtil>70?"var(--grn)":"var(--ora)"}}>{overallUtil}%</div><div style={{fontSize:11,color:"var(--muted)",fontWeight:600}}>Utilization</div></div>}/>
        <AGrid cols={4}><AKpi icon="📦" val={(totalAppr/1000).toFixed(1)+"K"} label="Approved Qty MT" trend="up" trendTxt="Total" color="var(--gold)"/><AKpi icon="✅" val={(totalSold/1000).toFixed(1)+"K"} label="Invoiced Qty MT" trend="up" trendTxt={overallUtil+"% used"} color="var(--grn)"/><AKpi icon="⏳" val={((totalAppr-totalSold)/1000).toFixed(1)+"K"} label="Pending Gap MT" trend="neu" trendTxt="Unsold" color="var(--ora)"/><AKpi icon="⚠️" val={lowUtil.length} label="Low Util lt50%" trend="down" trendTxt="Need follow-up" color="var(--red)"/></AGrid>
        <AGrid cols={2}>
          <ACard title="Grade-wise Utilization %" sub="Sold vs approved"><ABar items={gradeItems} colorFn={function(it){return Number(it.val)>70?"var(--grn)":Number(it.val)>40?"var(--gold)":"var(--red)";}}/></ACard>
          <ACard title="Low Utilization Orders" sub="Less than 50% invoiced">
            <div className="tw"><table><thead><tr><th>ID</th><th>Customer</th><th>Approved</th><th>Sold</th><th>Util%</th></tr></thead>
            <tbody>{lowUtil.map(r=><tr key={r.id}><td style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:"var(--gold-dk)",fontSize:11}}>{r.id}</td><td style={{fontWeight:600,fontSize:12}}>{r.customerName}</td><td>{r.qty} MT</td><td style={{color:"var(--grn)",fontWeight:700}}>{r.sold} MT</td><td><span style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:"var(--red)"}}>{r.util}%</span></td></tr>)}</tbody>
            </table></div>
          </ACard>
        </AGrid>
      </div>
    );
  }

  function AnCustomer() {
    const custList=customers||[];
    const custData=custList.map(function(c){const mine=ur.filter(r=>r.customerCode===c.customerCode);const app=mine.filter(r=>r.status==="Approved");return{...c,requests:mine.length,approved:app.length,vol:aSum(app,"qty"),avgNCR:aAvg(mine,"ncrPmt"),avgDisc:aAvg(mine,"difference")};}).filter(c=>c.requests>0);
    const topVol=[...custData].sort((a,b)=>b.vol-a.vol).slice(0,6);
    const topDisc=[...custData].sort((a,b)=>a.avgDisc-b.avgDisc).slice(0,5);
    const volItems=topVol.map(c=>({label:(c.customerName||c.customerCode).split(" ").slice(0,2).join(" "),val:c.vol,sub:c.vol+" MT"}));
    return (
      <div>
        <APh title="Customer Intelligence Dashboard" sub="Risk · Volume · Discount per customer"/>
        <AGrid cols={4}><AKpi icon="🏢" val={custList.length} label="Total Customers" trend="neu" trendTxt="In database" color="var(--gold)"/><AKpi icon="🔒" val={custData.filter(c=>c.approved===0&&c.requests>0).length} label="Zero Approval" trend="down" trendTxt="Risk accounts" color="var(--red)"/><AKpi icon="📊" val={custData.filter(c=>c.avgNCR<0).length} label="Neg NCR Customers" trend="down" trendTxt="Loss accounts" color="var(--ora)"/><AKpi icon="🌟" val={custData.filter(c=>c.vol>500).length} label="High Volume" trend="up" trendTxt=">500 MT" color="var(--grn)"/></AGrid>
        <AGrid cols={2}>
          <ACard title="Top Volume Customers" sub="By approved MT"><ABar items={volItems} colorFn={function(_,i){return["var(--gold)","#2563EB","var(--grn)","#7C3AED","var(--ora)","var(--red)"][i]||"var(--gold)";}}/></ACard>
          <ACard title="High Discount Customers" sub="Most negative difference from trade price">
            <div className="tw"><table><thead><tr><th>Customer</th><th>Avg Discount</th><th>Requests</th><th>Avg NCR</th></tr></thead>
            <tbody>{topDisc.map(c=><tr key={c.customerCode}><td style={{fontWeight:700}}>{c.customerName||c.customerCode}</td><td style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:"var(--red)"}}>₹{c.avgDisc.toFixed(0)}</td><td>{c.requests}</td><td style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:c.avgNCR<0?"var(--red)":"var(--grn)"}}>₹{c.avgNCR}</td></tr>)}</tbody>
            </table></div>
          </ACard>
        </AGrid>
      </div>
    );
  }

  function AnGeo() {
    const regions=[...new Set(ur.map(r=>r.region).filter(Boolean))];
    const zones=[...new Set(ur.map(r=>r.zone).filter(Boolean))];
    const regItems=regions.map(r=>({label:r.length>14?r.slice(0,12)+"…":r,val:aSum(ur.filter(x=>x.region===r&&x.status==="Approved"),"qty"),sub:aSum(ur.filter(x=>x.region===r&&x.status==="Approved"),"qty")+" MT"})).sort((a,b)=>b.val-a.val);
    const zoneItems=zones.map(z=>({label:z,val:ur.filter(x=>x.zone===z).length,sub:ur.filter(x=>x.zone===z).length+" req"}));
    const plantItems=PLANTS_LIST.map(p=>({label:p.replace(" Plant",""),val:aSum(ur.filter(r=>r.supplyFrom===p&&r.status==="Approved"),"qty"),sub:aSum(ur.filter(r=>r.supplyFrom===p&&r.status==="Approved"),"qty")+" MT"}));
    return (
      <div>
        <APh title="Geographic Analytics" sub="Region · Zone · Plant supply distribution"/>
        <AGrid cols={2}>
          <ACard title="Region-wise Approved Volume MT"><ABar items={regItems} colorFn={function(_,i){return["var(--gold)","var(--grn)","#2563EB","#7C3AED","var(--ora)"][i]||"var(--gold)";}}/></ACard>
          <ACard title="Zone-wise Request Count"><ABar items={zoneItems} colorFn={function(_,i){return["var(--gold)","#2563EB","var(--grn)","#7C3AED"][i]||"var(--gold)";}}/></ACard>
        </AGrid>
        <ACard title="Supply Plant-wise Volume"><ABar items={plantItems} colorFn={function(_,i){return["var(--gold)","#2563EB","#7C3AED"][i]||"var(--gold)";}}/></ACard>
        <ACard title="Region-wise Summary Table" style={{marginTop:14}}>
          <div className="tw"><table><thead><tr><th>Region</th><th>Requests</th><th>Approved</th><th>Volume MT</th><th>Avg NCR/MT</th></tr></thead>
          <tbody>{regions.map(function(r){const mine=ur.filter(x=>x.region===r);const app=mine.filter(x=>x.status==="Approved");return(<tr key={r}><td style={{fontWeight:700}}>{r}</td><td>{mine.length}</td><td>{app.length}</td><td style={{fontFamily:"'Sora',sans-serif",fontWeight:700,color:"#2563EB"}}>{aSum(app,"qty")} MT</td><td style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:aAvg(mine,"ncrPmt")<0?"var(--red)":"var(--grn)"}}>₹{aAvg(mine,"ncrPmt")}</td></tr>);})}</tbody>
          </table></div>
        </ACard>
      </div>
    );
  }

  function AnFreight() {
    const costs=[{name:"Primary Freight",key:"primaryFreight",color:"var(--gold)"},{name:"Secondary Freight",key:"secondaryFreight",color:"#2563EB"},{name:"Cost of Production",key:"costOfProduction",color:"#7C3AED"},{name:"Packing",key:"packing",color:"var(--grn)"},{name:"OP Commission",key:"opCommission",color:"var(--ora)"}];
    const costData=costs.map(c=>({...c,avg:aAvg(ur,c.key)}));
    const total=costData.reduce((s,c)=>s+c.avg,0);
    const plantData=PLANTS_LIST.map(function(p){const items=ur.filter(r=>r.supplyFrom===p);return{plant:p.replace(" Plant",""),pf:aAvg(items,"primaryFreight"),sf:aAvg(items,"secondaryFreight"),cop:aAvg(items,"costOfProduction")};});
    return (
      <div>
        <APh title="Freight and Cost Analytics" sub="Cost breakdown · Leakage analysis · Plant comparison" right={<div><div style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:22,color:"var(--ink)"}}>₹{total.toFixed(0)}</div><div style={{fontSize:11,color:"var(--muted)",fontWeight:600}}>Avg Total Cost/MT</div></div>}/>
        <AGrid cols={5}>{costData.map(c=><AKpi key={c.key} icon="💸" val={"₹"+c.avg} label={c.name} trend="neu" trendTxt={aPct(c.avg,total||1)+"%"} color={c.color}/>)}</AGrid>
        <ACard title="Cost Leakage Analysis" sub="Which component drives highest cost share">
          {costData.map(c=>(
            <div key={c.key} style={{marginBottom:14}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}><span style={{fontWeight:700}}>{c.name}</span><div style={{display:"flex",gap:12}}><span style={{fontWeight:700}}>₹{c.avg}/MT</span><span style={{fontWeight:700,color:c.color}}>{aPct(c.avg,total||1)}%</span></div></div>
              <div style={{background:"var(--border)",borderRadius:5,height:10,overflow:"hidden"}}><div style={{height:"100%",width:aPct(c.avg,total||1)+"%",background:c.color,borderRadius:5}}/></div>
            </div>
          ))}
        </ACard>
        <ACard title="Plant-wise Cost Comparison" style={{marginTop:14}}>
          <div className="tw"><table><thead><tr><th>Plant</th><th>Avg Primary Freight</th><th>Avg Secondary Freight</th><th>Avg COP</th></tr></thead>
          <tbody>{plantData.map(p=><tr key={p.plant}><td style={{fontWeight:700}}>{p.plant}</td><td style={{fontFamily:"'Sora',sans-serif",fontWeight:700,color:"var(--gold)"}}>₹{p.pf}</td><td style={{fontFamily:"'Sora',sans-serif",fontWeight:700,color:"#2563EB"}}>₹{p.sf}</td><td style={{fontFamily:"'Sora',sans-serif",fontWeight:700,color:"#7C3AED"}}>₹{p.cop}</td></tr>)}</tbody>
          </table></div>
        </ACard>
      </div>
    );
  }

  function AnAging() {
    const pnd=ur.filter(r=>r.status==="Pending");
    function getTat(r){try{return Math.floor((new Date()-new Date(r.date))/(86400000));}catch(e){return 0;}}
    const groups=[{label:"0-1 Day",min:0,max:1,color:"var(--grn)"},{label:"2-3 Days",min:2,max:3,color:"var(--gold)"},{label:"4-7 Days",min:4,max:7,color:"var(--ora)"},{label:">7 Days",min:8,max:9999,color:"var(--red)"}];
    const aged=groups.map(g=>({...g,count:pnd.filter(r=>{const t=getTat(r);return t>=g.min&&t<=g.max;}).length}));
    const critical=pnd.filter(r=>getTat(r)>5).slice(0,8);
    const LEVELS=["Sales Officer","Area Sales Manager","Regional Head","Zonal Head","Sales & Accounts"];
    const SHORT=["SO","ASM","RH","ZH","S&A"];
    const levelAging=LEVELS.map(function(lvl,i){return{name:SHORT[i],stuck:pnd.filter(r=>r.currentLevel===lvl&&getTat(r)>3).length};});
    return (
      <div>
        <APh title="Approval Aging Dashboard" sub="Pending buckets · Escalation risk" right={<div style={{padding:"8px 14px",background:critical.length>0?"var(--red-f)":"var(--grn-f)",borderRadius:10,border:"1.5px solid",borderColor:critical.length>0?"#FCA5A5":"#A7F3D0"}}><div style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:22,color:critical.length>0?"var(--red)":"var(--grn)"}}>{critical.length}</div><div style={{fontSize:11,fontWeight:600,color:"var(--muted)"}}>Critical</div></div>}/>
        <AGrid cols={4}>{aged.map(a=><AKpi key={a.label} icon="📋" val={a.count} label={a.label} trend={a.min>3?"down":"up"} trendTxt="Pending" color={a.color}/>)}</AGrid>
        <AGrid cols={2}>
          <ACard title="Aging Bucket Bars"><ABar items={aged.map(a=>({label:a.label,val:a.count,sub:a.count+" req"}))} colorFn={function(_,i){return aged[i].color;}}/></ACard>
          <ACard title="Escalation Risk by Level" sub="Requests stuck more than 3 days">
            {levelAging.map(l=>(
              <div key={l.name} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 0",borderBottom:"1px solid var(--border2)"}}>
                <div style={{width:40,fontSize:12,fontWeight:700}}>{l.name}</div>
                <div style={{flex:1,background:"var(--border)",borderRadius:4,height:14,overflow:"hidden"}}><div style={{height:"100%",width:Math.max(5,aPct(l.stuck,Math.max(...levelAging.map(x=>x.stuck),1)))+"%",background:l.stuck>2?"var(--red)":l.stuck>0?"var(--ora)":"var(--grn)",borderRadius:4}}/></div>
                <span style={{width:60,fontFamily:"'Sora',sans-serif",fontWeight:800,color:l.stuck>2?"var(--red)":l.stuck>0?"var(--ora)":"var(--grn)",textAlign:"right"}}>{l.stuck} stuck</span>
              </div>
            ))}
          </ACard>
        </AGrid>
        <ACard title="Critical Pending - More than 5 Days Old" sub="Immediate escalation required">
          <div className="tw"><table><thead><tr><th>Request ID</th><th>Customer</th><th>Grade</th><th>Level</th><th>Days Old</th><th>Priority</th></tr></thead>
          <tbody>{critical.length===0&&<tr><td colSpan={6} style={{textAlign:"center",padding:24,color:"var(--grn)",fontWeight:700}}>No critical requests</td></tr>}
          {critical.map(function(r){const t=getTat(r);return(<tr key={r.id}><td style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:"var(--gold-dk)",fontSize:12}}>{r.id}</td><td style={{fontWeight:600}}>{r.customerName}</td><td><span className="badge b-grade">{r.grade}</span></td><td style={{fontSize:12}}>{r.currentLevel}</td><td style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:t>7?"var(--red)":"var(--ora)"}}>{t}d</td><td><span className={t>7?"badge b-rejected":"badge b-pending"}>{t>7?"Critical":"High"}</span></td></tr>);})}</tbody>
          </table></div>
        </ACard>
      </div>
    );
  }

  function AnAudit() {
    const exceptions=ur.filter(r=>r.ncrPmt<-300&&r.status==="Approved");
    const excessDisc=ur.filter(r=>r.difference<-30&&r.status==="Approved");
    const blk=ur.filter(r=>r.blocked);
    const sentBack=ur.filter(r=>(r.history||[]).some(function(h){return h.action==="Sent Back";}));
    const compliance=Math.max(0,100-exceptions.length*5-excessDisc.length*2);
    return (
      <div>
        <APh title="Audit and Compliance Dashboard" sub="Policy exceptions · NCR threshold breaches · Full audit trail" right={<div style={{padding:"10px 16px",background:compliance>80?"var(--grn-f)":"var(--red-f)",borderRadius:10,border:"1.5px solid",borderColor:compliance>80?"#A7F3D0":"#FCA5A5"}}><div style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:22,color:compliance>80?"var(--grn)":"var(--red)"}}>{compliance}%</div><div style={{fontSize:11,fontWeight:600,color:"var(--muted)"}}>Compliance</div></div>}/>
        <AGrid cols={4}><AKpi icon="⚠️" val={exceptions.length} label="NCR Exceptions" trend="down" trendTxt="Approved lt-300" color="var(--red)"/><AKpi icon="💸" val={excessDisc.length} label="Excess Discount" trend="down" trendTxt="Discount gt 30" color="var(--ora)"/><AKpi icon="🔒" val={blk.length} label="Blocked Orders" trend="neu" trendTxt="Admin blocked" color="var(--gold)"/><AKpi icon="↩" val={sentBack.length} label="Sent Back" trend="neu" trendTxt="Revision requests" color="#2563EB"/></AGrid>
        <ACard title="Policy Exception Breakdown">
          {[{label:"NCR below -300 threshold",count:exceptions.length,c:"var(--red)"},{label:"Excess discount (more than 30)",count:excessDisc.length,c:"var(--ora)"},{label:"Blocked customer orders",count:blk.length,c:"var(--gold)"},{label:"Sent Back for Revision",count:sentBack.length,c:"#2563EB"}].map(function(e,i){return(
            <div key={i} style={{display:"flex",alignItems:"center",gap:12,padding:"10px 0",borderBottom:"1px solid var(--border2)"}}>
              <div style={{flex:1,fontSize:13,fontWeight:600}}>{e.label}</div>
              <span style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:18,color:e.c}}>{e.count}</span>
              <span className={e.count>3?"badge b-rejected":e.count>0?"badge b-pending":"badge b-approved"}>{e.count>3?"High":e.count>0?"Medium":"Clear"}</span>
            </div>
          );})}
        </ACard>
        <ACard title="Policy Exceptions - Approved with NCR below -300" style={{marginTop:14}}>
          <div className="tw"><table><thead><tr><th>Request ID</th><th>Customer</th><th>Grade</th><th>Region</th><th>NCR/MT</th></tr></thead>
          <tbody>{exceptions.length===0&&<tr><td colSpan={5} style={{textAlign:"center",padding:24,color:"var(--grn)",fontWeight:700}}>No exceptions found</td></tr>}
          {exceptions.slice(0,8).map(r=><tr key={r.id}><td style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:"var(--gold-dk)",fontSize:12}}>{r.id}</td><td style={{fontWeight:600}}>{r.customerName}</td><td><span className="badge b-grade">{r.grade}</span></td><td style={{fontSize:12}}>{r.region}</td><td style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:"var(--red)"}}>₹{Number(r.ncrPmt).toFixed(0)}</td></tr>)}</tbody>
          </table></div>
        </ACard>
      </div>
    );
  }

  function AnTrade() {
    const avgDiff=aAvg(ur,"difference");
    const regions=[...new Set(ur.map(r=>r.region).filter(Boolean))];
    const gradeData=GRADES_LIST.map(g=>({label:g,val:aAvg(ur.filter(r=>r.grade===g),"orderPrice"),sub:"₹"+aAvg(ur.filter(r=>r.grade===g),"orderPrice")}));
    const tradeData=GRADES_LIST.map(g=>({label:g,val:aAvg(ur.filter(r=>r.grade===g),"tradePrice"),sub:"₹"+aAvg(ur.filter(r=>r.grade===g),"tradePrice")}));
    const regDisc=regions.map(r=>({label:r.length>12?r.slice(0,10)+"…":r,val:Math.abs(aAvg(ur.filter(x=>x.region===r),"difference")),sub:"₹"+aAvg(ur.filter(x=>x.region===r),"difference")}));
    return (
      <div>
        <APh title="Trade Price vs Non-Trade Analysis" sub="Discount tracking · Market competitiveness · Price realization" right={<div><div style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:24,color:avgDiff<0?"var(--red)":"var(--grn)"}}>₹{avgDiff.toFixed(0)}</div><div style={{fontSize:11,color:"var(--muted)",fontWeight:600}}>Avg Discount/MT</div></div>}/>
        <AGrid cols={4}><AKpi icon="💰" val={"₹"+aAvg(ur,"tradePrice").toFixed(0)} label="Avg Trade Price" trend="neu" trendTxt="Reference" color="var(--gold)"/><AKpi icon="📋" val={"₹"+aAvg(ur,"orderPrice").toFixed(0)} label="Avg Order Price" trend="neu" trendTxt="Deal price" color="#2563EB"/><AKpi icon="📉" val={"₹"+Math.abs(avgDiff).toFixed(0)} label="Avg Discount" trend="down" trendTxt="Below trade" color="var(--red)"/><AKpi icon="🎯" val={aPct(ur.filter(r=>r.difference>=0).length,ur.length)+"%"} label="Above Trade Price" trend="up" trendTxt="Premium deals" color="var(--grn)"/></AGrid>
        <AGrid cols={2}>
          <ACard title="Grade-wise Order Price"><ABar items={gradeData} colorFn={function(_,i){return["var(--gold)","#2563EB","#7C3AED"][i]||"var(--gold)";}}/></ACard>
          <ACard title="Region Competitiveness - Avg Discount" sub="Negative = below trade price">
            <ABar items={regDisc} colorFn={function(it){return Number(it.sub.replace("₹",""))<-15?"var(--red)":Number(it.sub.replace("₹",""))<0?"var(--ora)":"var(--grn)";}}/>
          </ACard>
        </AGrid>
      </div>
    );
  }

  function AnMaster() {
    const plantVol=PLANTS_LIST.map(p=>({label:p.replace(" Plant",""),val:aSum(ur.filter(r=>r.supplyFrom===p&&r.status==="Approved"),"qty"),sub:ur.filter(r=>r.supplyFrom===p).length+" requests"}));
    const regions=[...new Set(ur.map(r=>r.region).filter(Boolean))];
    const regCust=regions.map(r=>({label:r.length>12?r.slice(0,10)+"…":r,val:ur.filter(x=>x.region===r).length,sub:ur.filter(x=>x.region===r).length+" req"}));
    return (
      <div>
        <APh title="Master Data Dashboard" sub="Customer · Plant · Supply source analytics"/>
        <AGrid cols={4}><AKpi icon="🏢" val={(customers||[]).length} label="Total Customers" trend="neu" trendTxt="In database" color="var(--gold)"/><AKpi icon="✅" val={(customers||[]).filter(c=>c.status==="Active"||!c.status).length} label="Active Customers" trend="up" trendTxt="Live accounts" color="var(--grn)"/><AKpi icon="🏗" val={PLANTS_LIST.length} label="Plants and Depots" trend="neu" trendTxt="Supply sources" color="#2563EB"/><AKpi icon="📦" val={(aSum(ur.filter(r=>r.status==="Approved"),"qty")/1000).toFixed(1)+"K MT"} label="Total Dispatched" trend="up" trendTxt="MT approved" color="#7C3AED"/></AGrid>
        <AGrid cols={2}>
          <ACard title="Plant-wise Volume Dispatched"><ABar items={plantVol} colorFn={function(_,i){return["var(--gold)","#2563EB","#7C3AED"][i]||"var(--gold)";}}/></ACard>
          <ACard title="Region-wise Request Count"><ABar items={regCust} colorFn={function(_,i){return["var(--gold)","var(--grn)","#2563EB","#7C3AED"][i]||"var(--gold)";}}/></ACard>
        </AGrid>
      </div>
    );
  }

  function AnForecast() {
    const pipeline=ur.filter(r=>r.status==="Pending");
    const approved=ur.filter(r=>r.status==="Approved");
    const regions=[...new Set(ur.map(r=>r.region).filter(Boolean))];
    const plantPlan=PLANTS_LIST.map(p=>({name:p.replace(" Plant",""),pending:aSum(ur.filter(r=>r.supplyFrom===p&&r.status==="Pending"),"qty"),approved:aSum(ur.filter(r=>r.supplyFrom===p&&r.status==="Approved"),"qty")}));
    const regForecast=regions.map(r=>({label:r.length>12?r.slice(0,10)+"…":r,val:aSum(ur.filter(x=>x.region===r&&x.status==="Pending"),"qty"),sub:aSum(ur.filter(x=>x.region===r&&x.status==="Pending"),"qty")+" MT pipeline"}));
    return (
      <div>
        <APh title="Sales Forecast and Planning" sub="Pipeline · Demand forecast · Plant dispatch planning"/>
        <AGrid cols={4}><AKpi icon="🔮" val={(aSum(pipeline,"qty")/1000).toFixed(1)+"K"} label="Pipeline Volume" trend="up" trendTxt="MT pending approval" color="var(--gold)"/><AKpi icon="📊" val={pipeline.length} label="Active Pipeline" trend="up" trendTxt="Pending requests" color="#2563EB"/><AKpi icon="🏭" val={(aSum(approved,"qty")/1000).toFixed(1)+"K"} label="Approved Volume" trend="up" trendTxt="MT total" color="var(--grn)"/><AKpi icon="🎯" val={aAvg(approved,"qty").toFixed(0)+" MT"} label="Avg Deal Size" trend="neu" trendTxt="Per approval" color="#7C3AED"/></AGrid>
        <AGrid cols={2}>
          <ACard title="Region Demand Pipeline" sub="Pending volume by region"><ABar items={regForecast} colorFn={function(_,i){return["var(--gold)","#2563EB","var(--grn)","#7C3AED"][i]||"var(--gold)";}}/></ACard>
          <ACard title="Plant Dispatch Planning" sub="Pipeline vs approved per plant">
            {plantPlan.map(function(p,i){return(
              <div key={p.name} style={{marginBottom:18}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}><span style={{fontWeight:700,fontSize:14}}>{p.name}</span><div style={{display:"flex",gap:16}}><div><div style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:"var(--ora)"}}>{p.pending} MT</div><div style={{fontSize:10,color:"var(--muted)"}}>PIPELINE</div></div><div><div style={{fontFamily:"'Sora',sans-serif",fontWeight:800,color:"var(--gold)"}}>{p.approved} MT</div><div style={{fontSize:10,color:"var(--muted)"}}>APPROVED</div></div></div></div>
                <div style={{background:"var(--border)",borderRadius:6,height:12,overflow:"hidden"}}><div style={{height:"100%",width:aPct(p.pending,Math.max(p.approved,p.pending,1))+"%",background:["var(--gold)","#2563EB","#7C3AED"][i]||"var(--gold)",borderRadius:6}}/></div>
              </div>
            );})}
          </ACard>
        </AGrid>
      </div>
    );
  }

  function renderAnalytics(key) {
    if (key==="exec")     return <AnExec/>;
    if (key==="workflow") return <AnWorkflow/>;
    if (key==="ncr")      return <AnNCR/>;
    if (key==="perf")     return <AnPerf/>;
    if (key==="sales")    return <AnSales/>;
    if (key==="customer") return <AnCustomer/>;
    if (key==="geo")      return <AnGeo/>;
    if (key==="freight")  return <AnFreight/>;
    if (key==="aging")    return <AnAging/>;
    if (key==="audit")    return <AnAudit/>;
    if (key==="trade")    return <AnTrade/>;
    if (key==="master")   return <AnMaster/>;
    if (key==="forecast") return <AnForecast/>;
    return null;
  }
  // existing Dashboard state/calculations
  const ur2=ur; // already computed above
  const pending   = ur.filter(r => r.status === "Pending");
  const approved  = ur.filter(r => r.status === "Approved");
  const rejected  = ur.filter(r => r.status === "Rejected");
  const blocked   = ur.filter(r => r.blocked === true);
  const completed = ur.filter(r => r.status === "Approved" && r.validityTo && r.validityTo < TODAY);
  const ncrCritical = ur.filter(r => r.ncrPmt < -300);
  const totalVol  = ur.reduce((s, r) => s + Number(r.qty || 0), 0);
  const approvalRate = ur.length ? Math.round(approved.length / ur.length * 100) : 0;
  const gradeData = [{ g: "OPC", n: ur.filter(r => r.grade === "OPC").length }, { g: "PPC", n: ur.filter(r => r.grade === "PPC").length }, { g: "ProMaxX", n: ur.filter(r => r.grade === "ProMaxX").length }];
  const maxGrade = Math.max(...gradeData.map(g => g.n), 1);
  const regionData = [...new Set(ur.map(r => r.region))].map(reg => ({ reg, n: ur.filter(r => r.region === reg).length })).sort((a, b) => b.n - a.n).slice(0, 5);
  const maxRegion = Math.max(...regionData.map(r => r.n), 1);
  const mine = ur.filter(r => r.currentLevel === user.role && r.status === "Pending");
  const GRADE_COLORS = { "OPC": "linear-gradient(90deg,var(--gold-lt),var(--gold))", "PPC": "linear-gradient(90deg,#38BDF8,#0EA5E9)", "ProMaxX": "linear-gradient(90deg,#A78BFA,#7C3AED)" };
  const REGION_COLORS = ["var(--gold)", "#0EA5E9", "#10B981", "#8B5CF6", "#F59E0B"];
  const recentActivity = ur.flatMap(r => r.history.map(h => ({ ...h, reqId: r.id, custName: r.customerName }))).sort((a, b) => b.time.localeCompare(a.time)).slice(0, 6);
  const ncrByGrade = ["OPC", "PPC", "ProMaxX"].map(g => { const gR = ur.filter(r => r.grade === g); const avg = gR.length ? gR.reduce((s, r) => s + Number(r.ncrPmt || 0), 0) / gR.length : 0; return { g, avg: avg.toFixed(0), count: gR.length }; });

  return (
    <div>
      <div>
      <div className="ph">
        <div><div className="ph-eyebrow">Overview</div><div className="ph-title">Dashboard</div><div className="ph-sub">Welcome back, <strong style={{ color: "var(--gold-dk)" }}>{user.name}</strong> · {user.role}</div></div>
        {(user.role === "Sales Officer" || user.role === "Admin") && <button className="btn btn-primary" onClick={() => setPage("new-request")}>+ New Price Request</button>}
      </div>
      {mine.length > 0 && (
        <div className="alert a-warning" style={{ marginBottom: 20 }}>
          <span style={{ fontSize: 20 }}>🔔</span>
          <div><strong>{mine.length} request{mine.length > 1 ? "s" : ""} pending your action.</strong></div>
          <button className="btn btn-outline btn-sm" style={{ marginLeft: "auto", flexShrink: 0 }} onClick={() => { setPage("requests"); setActiveTab("pending"); }}>View →</button>
        </div>
      )}
      {blocked.length > 0 && (
        <div className="alert a-danger" style={{ marginBottom: 20, cursor: "pointer" }} onClick={() => { setPage("requests"); setActiveTab("blocked"); }}>
          <span style={{ fontSize: 20 }}>🔒</span>
          <div><strong>{blocked.length} order{blocked.length > 1 ? "s" : ""} currently blocked by Admin.</strong> <span style={{ fontWeight: 500 }}>Click to view blocked orders.</span></div>
          <span style={{ marginLeft: "auto", fontSize: 12, fontWeight: 700, color: "#991B1B" }}>View →</span>
        </div>
      )}
      <div className="kpi-grid">
        {[
          { label: "Total Requests", val: ur.length, icon: "📋", color: "var(--gold)", bg: "var(--gold-f)", trend: "neutral", trendTxt: "All time" },
          { label: "Pending", val: pending.length, icon: "🕐", color: "#F59E0B", bg: "#FFF8E1", trend: pending.length > 2 ? "down" : "up", trendTxt: "Needs action", onClick: () => { setPage("requests"); setActiveTab("pending"); } },
          { label: "Approved", val: approved.length, icon: "✅", color: "#059669", bg: "var(--grn-f)", trend: "up", trendTxt: `${approvalRate}% rate`, onClick: () => { setPage("requests"); setActiveTab("approved"); } },
          { label: "Rejected", val: rejected.length, icon: "❌", color: "#DC2626", bg: "var(--red-f)", trend: "neutral", trendTxt: "This period", onClick: () => { setPage("requests"); setActiveTab("rejected"); } },
          { label: "🔒 Blocked Orders", val: blocked.length, icon: "🔒", color: "#C2410C", bg: "#FFF7ED", trend: blocked.length > 0 ? "down" : "up", trendTxt: "Admin blocked", onClick: () => { setPage("requests"); setActiveTab("blocked"); } },
          { label: "Completed (Expired)", val: completed.length, icon: "🏁", color: "#0369A1", bg: "#F0F9FF", trend: "neutral", trendTxt: "Validity over", onClick: () => { setPage("requests"); setActiveTab("completed"); } },
          { label: "Total Volume", val: `${totalVol.toLocaleString()}`, icon: "📦", color: "#7C3AED", bg: "#F5F3FF", trend: "up", trendTxt: "MT ordered" },
          { label: "NCR Alerts", val: ncrCritical.length, icon: "⚠️", color: "#EA580C", bg: "#FFF7ED", trend: ncrCritical.length > 0 ? "down" : "up", trendTxt: "NCR < -300" },
        ].map((k, i) => (
          <div key={i} className="kpi" onClick={k.onClick} style={{ cursor: k.onClick ? "pointer" : "default" }}>
            <div className="kpi-accent" style={{ background: k.color }} />
            <div className="kpi-icon" style={{ background: k.bg }}>{k.icon}</div>
            <div className="kpi-val">{k.val}</div>
            <div className="kpi-label">{k.label}</div>
            <div className={`kpi-trend ${k.trend}`}>{k.trend === "up" ? "↑" : k.trend === "down" ? "↓" : "→"} {k.trendTxt}</div>
          </div>
        ))}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">Volume by Grade</div><div className="card-sub">Request distribution across grades</div></div></div>
          <div className="card-body">
            <div className="bar-chart">
              {gradeData.map(g => (
                <div key={g.g} className="bar-row">
                  <div className="bar-label">{g.g}</div>
                  <div className="bar-track"><div className="bar-fill" style={{ width: `${(g.n / maxGrade) * 100}%`, background: GRADE_COLORS[g.g] }}>{g.n > 0 ? `${g.n} req` : ""}</div></div>
                  <div className="bar-val">{g.n}</div>
                </div>
              ))}
            </div>
            <div style={{ marginTop: 20, borderTop: "1.5px solid var(--border2)", paddingTop: 16 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "var(--muted)", textTransform: "uppercase", letterSpacing: .5, marginBottom: 10 }}>Avg NCR / MT by Grade</div>
              {ncrByGrade.map(g => (
                <div key={g.g} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: "1px solid var(--border2)" }}>
                  <span style={{ fontSize: 13, fontWeight: 600, color: "var(--ink3)" }}>{g.g} <span style={{ color: "var(--muted)", fontWeight: 500 }}>({g.count})</span></span>
                  <span style={{ fontFamily: "'Sora',sans-serif", fontSize: 16, fontWeight: 800, color: Number(g.avg) < -300 ? "var(--red)" : "var(--grn)" }}>₹{g.avg}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">Requests by Region</div><div className="card-sub">Geographic distribution</div></div></div>
          <div className="card-body">
            <div className="bar-chart">
              {regionData.map((r, i) => (
                <div key={r.reg} className="bar-row">
                  <div className="bar-label" style={{ fontSize: 11.5 }}>{r.reg}</div>
                  <div className="bar-track"><div className="bar-fill" style={{ width: `${(r.n / maxRegion) * 100}%`, background: REGION_COLORS[i] || "var(--gold)" }}>{r.n > 0 ? r.n : ""}</div></div>
                  <div className="bar-val">{r.n}</div>
                </div>
              ))}
            </div>
            <div style={{ marginTop: 20, borderTop: "1.5px solid var(--border2)", paddingTop: 16 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "var(--muted)", textTransform: "uppercase", letterSpacing: .5, marginBottom: 12 }}>Status Breakdown</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
                {[{ l: "Pending", n: pending.length, c: "#F59E0B", bg: "#FFF8E1" }, { l: "Approved", n: approved.length, c: "var(--grn)", bg: "var(--grn-f)" }, { l: "Rejected", n: rejected.length, c: "var(--red)", bg: "var(--red-f)" }].map(s => (
                  <div key={s.l} style={{ background: s.bg, borderRadius: 10, padding: "12px 10px", textAlign: "center", border: "1.5px solid", borderColor: s.c + "40" }}>
                    <div style={{ fontFamily: "'Sora',sans-serif", fontSize: 22, fontWeight: 800, color: s.c }}>{s.n}</div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "var(--muted)", textTransform: "uppercase", letterSpacing: .5, marginTop: 4 }}>{s.l}</div>
                    <div style={{ marginTop: 6 }}><div className="prog-bar"><div className="prog-fill" style={{ width: `${ur.length ? s.n / ur.length * 100 : 0}%`, background: s.c }} /></div></div>
                    <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 4, fontWeight: 600 }}>{ur.length ? Math.round(s.n / ur.length * 100) : 0}%</div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">Recent Requests</div></div><button className="btn btn-ghost btn-sm" onClick={() => setPage("requests")}>View all →</button></div>
          <div className="tw">
            <table>
              <thead><tr><th>ID</th><th>Customer</th><th>Grade</th><th>NCR</th><th>Status</th></tr></thead>
              <tbody>{ur.slice(0, 6).map(r => (
                <tr key={r.id}>
                  <td><span style={{ fontFamily: "'Sora',sans-serif", fontWeight: 700, color: "var(--gold-dk)", fontSize: 12 }}>{r.id}</span></td>
                  <td><div style={{ fontWeight: 600, fontSize: 13 }}>{r.customerName.split(" ").slice(0, 2).join(" ")}</div></td>
                  <td><span className="badge b-grade">{r.grade}</span></td>
                  <td><span style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: r.ncrPmt < -300 ? "var(--red)" : "var(--grn)", fontSize: 14 }}>₹{Number(r.ncrPmt).toFixed(0)}</span></td>
                  <td><StatusBadge status={r.status} /></td>
                </tr>
              ))}</tbody>
            </table>
          </div>
        </div>
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">Activity Feed</div><div className="card-sub">Recent actions</div></div></div>
          <div className="card-body" style={{ padding: "8px 20px" }}>
            <div className="feed">
              {recentActivity.map((h, i) => {
                const actColor = { Created: "var(--blu-f)", Approved: "var(--grn-f)", Rejected: "var(--red-f)", Validated: "var(--gold-f)" }[h.action] || "var(--cream)";
                const actIcon = { Created: "📝", Approved: "✅", Rejected: "❌", Validated: "🔍", Modified: "✏️" }[h.action] || "•";
                return (
                  <div key={i} className="feed-item">
                    <div className="feed-dot" style={{ background: actColor }}>{actIcon}</div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div className="feed-title">{h.action} · <span style={{ color: "var(--gold-dk)" }}>{h.reqId}</span></div>
                      <div className="feed-sub">{h.custName} — {h.level}</div>
                      <div className="feed-time">{h.time}</div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>
      </div>
    </div>
  );
}

// ═══ REQUESTS PAGE
function RequestsPage({ requests, user, users, freightMaster, salesEntries, setSalesEntries, onAction, activeTab, setActiveTab }) {
  const [sel, setSel] = useState(null);
  const { exportCsv, ExportModal, ToastEl } = useExport();
  const TABS = [
    { key: "all", label: "All" },
    { key: "pending", label: `⏳ Pending (${requests.filter(r=>r.status==="Pending").length})` },
    { key: "approved", label: `✅ Approved` },
    { key: "rejected", label: "❌ Rejected" },
    { key: "blocked", label: `🔒 Blocked (${requests.filter(r=>r.blocked).length})` },
    { key: "completed", label: `🏁 Completed` },
    { key: "ncr-opc", label: "⚠ NCR Alert" },
  ];
  const filtered = requests.filter(r => {
    if (activeTab === "pending" && r.status !== "Pending") return false;
    if (activeTab === "approved" && (r.status !== "Approved" || r.blocked)) return false;
    if (activeTab === "rejected" && r.status !== "Rejected") return false;
    if (activeTab === "blocked" && !r.blocked) return false;
    if (activeTab === "completed" && !(r.status === "Approved" && r.validityTo && r.validityTo < TODAY)) return false;
    if (activeTab === "ncr-opc" && !(r.ncrPmt < -300)) return false;
    if (user.role !== "Admin" && user.role !== "Sales & Accounts" && user.role !== "Zonal Head") { if (!user.assignedCustomers?.includes(r.customerCode)) return false; }
    return true;
  });
  return (
    <div>
      {ToastEl}
      <div className="ph">
        <div><div className="ph-eyebrow">Workflow</div><div className="ph-title">Price Requests</div><div className="ph-sub">{filtered.length} records</div></div>
        <button className="btn btn-secondary btn-sm" onClick={() => exportCsv(filtered, REQUEST_REPORT_COLS, `price_requests_${new Date().toISOString().split("T")[0]}.csv`, "Price Requests")}>⬇ Export</button>
      </div>
      {activeTab === "blocked" && filtered.length > 0 && (
        <div className="alert a-danger" style={{ marginBottom: 16 }}>
          <span>🔒</span>
          <div><strong>{filtered.length} order(s) are currently blocked.</strong> Blocked orders cannot be dispatched. Admin can unblock them below.</div>
        </div>
      )}
      {activeTab === "completed" && (
        <div className="alert a-info" style={{ marginBottom: 16 }}>
          <span>🏁</span>
          <div><strong>Completed Orders</strong> — These approved orders have passed their validity date and are no longer active for dispatch.</div>
        </div>
      )}
      <div className="tabs">{TABS.map(t => <div key={t.key} className={`tab ${activeTab === t.key ? "active" : ""}`} onClick={() => setActiveTab(t.key)}>{t.label}</div>)}</div>
      <div className="card">
        <div className="tw"><table>
          <thead><tr><th>Request ID</th><th>Date</th><th>Customer</th><th>Region</th><th>Grade</th><th>Qty</th><th>Order Price</th><th>NCR/MT</th><th>Validity</th><th>Status</th><th>Level</th><th>Action</th></tr></thead>
          <tbody>
            {filtered.length === 0 && <tr><td colSpan={12} style={{ textAlign: "center", padding: "48px", color: "var(--muted)" }}><div style={{ fontSize: 40, marginBottom: 12 }}>📋</div><div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 700, fontSize: 16 }}>No requests</div></td></tr>}
            {filtered.map(r => (
              <tr key={r.id} style={{ cursor: "pointer", background: r.blocked ? "#FFF7ED" : undefined }} onClick={() => setSel(r)}>
                <td><span style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: "var(--gold-dk)", fontSize: 12.5 }}>{r.id}</span></td>
                <td style={{ color: "var(--muted)", fontSize: 12, fontWeight: 600 }}>{r.date}</td>
                <td><div style={{ fontWeight: 700, fontSize: 13 }}>{r.customerName}</div><div style={{ fontSize: 11, color: "var(--muted)" }}>{r.customerCode}</div></td>
                <td style={{ fontSize: 12.5, fontWeight: 600 }}>{r.region}</td>
                <td><span className="badge b-grade">{r.grade}/{r.materialType}</span></td>
                <td style={{ fontWeight: 700 }}>{r.qty}</td>
                <td style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 14 }}>₹{r.orderPrice}</td>
                <td><span style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 15, color: r.ncrPmt < -300 ? "var(--red)" : "var(--grn)" }}>₹{Number(r.ncrPmt).toFixed(0)}</span></td>
                <td style={{ fontSize: 11.5, color: "var(--muted)", fontWeight: 500 }}>{r.validityFrom} → {r.validityTo}</td>
                <td><StatusBadge status={r.status} blocked={r.blocked} /></td>
                <td style={{ fontSize: 11.5, color: "var(--muted)", fontWeight: 600 }}>{r.currentLevel}</td>
                <td onClick={e => e.stopPropagation()}>
                  <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                    {/* Sales Officer: only shown edit-resubmit button when request is sent back to them */}
                    {user.role === "Sales Officer" && isSentBack(r) && (
                      <button className="btn btn-xs btn-warning" style={{background:"#FEF3C7",color:"#92400E",border:"1px solid #FCD34D"}} onClick={() => setSel(r)}>✏️ Edit & Resubmit</button>
                    )}
                    {canAct(user, r) && user.role !== "Sales Officer" && (
                      <>
                        <button className="btn btn-xs btn-success" onClick={() => onAction(r.id, "approve", "", r.orderPrice, r.qty)}>✓</button>
                        <button className="btn btn-xs btn-secondary" onClick={() => setSel(r)}>Detail</button>
                      </>
                    )}
                    {user.role === "Admin" && r.status === "Approved" && !r.blocked && (
                      <button className="btn btn-xs" style={{ background: "#FFF0E0", color: "#C2410C", border: "1px solid #FDBA74" }} onClick={() => onAction(r.id, "block", "Admin blocked", r.orderPrice, r.qty)}>🔒 Block</button>
                    )}
                    {user.role === "Admin" && r.blocked && (
                      <button className="btn btn-xs btn-success" onClick={() => onAction(r.id, "unblock", "Admin unblocked", r.orderPrice, r.qty)}>🔓 Unblock</button>
                    )}
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table></div>
      </div>
      {sel && <ReqDetailModal req={sel} user={user} users={users} freightMaster={freightMaster} salesEntries={salesEntries} setSalesEntries={setSalesEntries} onClose={() => setSel(null)} onAction={onAction} />}
      {ExportModal}
    </div>
  );
}

function ReqDetailModal({ req, user, users, freightMaster, salesEntries, setSalesEntries, onClose, onAction }) {
  const [tab, setTab] = useState("details");
  const [act, setAct] = useState(null);
  const [remark, setRemark] = useState("");
  const [mp, setMp] = useState(req.orderPrice);
  const [mq, setMq] = useState(req.qty);
  const [mValidFrom, setMValidFrom] = useState(req.validityFrom || "");
  const [mValidTo, setMValidTo]     = useState(req.validityTo   || "");

  // Resubmit edit fields — pre-filled from existing request
  const [showResubmit, setShowResubmit] = useState(false);
  const [rsForm, setRsForm] = useState({
    orderPrice:  req.orderPrice,
    qty:         req.qty,
    validityFrom:req.validityFrom || "",
    validityTo:  req.validityTo   || "",
    paymentTerms:req.paymentTerms || "Advance",
    remark:      "",
  });

  const myEntries = (salesEntries||[]).filter(e=>e.requestId===req.id);
  const canDo = canAct(user, req) && user.role !== "Sales Officer";
  // SO can edit & resubmit if request was sent back to Sales Officer level
  const soCanResubmit = user.role === "Sales Officer" && isSentBack(req);

  const dist = (freightMaster||[]).find(f => f.to === req.destination || f.route === req.destination);
  const doAction = t => { if (t === "reject" && !remark.trim()) return; if (t === "sendback" && !remark.trim()) return; onAction(req.id, t, remark, mp, mq, mValidFrom, mValidTo); setAct(null); onClose(); };

  const doResubmit = () => {
    if (!rsForm.remark.trim()) return;
    onAction(req.id, "resubmit", rsForm.remark, rsForm.orderPrice, rsForm.qty, rsForm.validityFrom, rsForm.validityTo, rsForm.paymentTerms);
    onClose();
  };
  const F = ({ label, value }) => (<div><div className="dv-lbl">{label}</div><div className={`dv-val${!value ? " empty" : ""}`}>{value || "—"}</div></div>);
  const FILE_ICONS = { pdf: "📄", doc: "📝", docx: "📝", xls: "📊", xlsx: "📊", png: "🖼️", jpg: "🖼️", jpeg: "🖼️", gif: "🖼️", txt: "📃", zip: "🗜️", rar: "🗜️" };
  const getFileIcon = name => { const ext = (name||"").split(".").pop().toLowerCase(); return FILE_ICONS[ext] || "📎"; };
  const getFileColor = name => { const ext = (name||"").split(".").pop().toLowerCase(); if(["pdf"].includes(ext)) return "#FEF2F2"; if(["xls","xlsx"].includes(ext)) return "#ECFDF5"; if(["png","jpg","jpeg","gif"].includes(ext)) return "#EFF6FF"; if(["doc","docx"].includes(ext)) return "#EFF6FF"; return "var(--cream)"; };
  const fmtSize = b => b < 1024 ? `${b} B` : b < 1048576 ? `${(b/1024).toFixed(1)} KB` : `${(b/1048576).toFixed(1)} MB`;
  const downloadAtt = att => { const a = document.createElement("a"); a.href = att.dataUrl; a.download = att.name; a.click(); };
  const atts = req.attachments || [];
  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal modal-xl">
        <div className="m-hdr">
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}><div className="m-hdr-accent">📋</div><div><div className="m-title">{req.id}</div><div className="m-sub">{req.customerName} · {req.date}</div></div></div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}><StatusBadge status={req.status} blocked={req.blocked} /><span style={{fontSize:12,color:"var(--muted)",padding:"4px 8px",borderRadius:6,background:"var(--cream)",border:"1px solid var(--border2)"}}>{myEntries.length>0?`📊 ${myEntries.length} invoices recorded`:"📝 Add in Sale Updation"}</span><button className="m-close" onClick={onClose}>×</button></div>
        </div>
        <div className="m-body">
          <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 16, padding: "12px 16px", background: "var(--cream)", borderRadius: 10, border: "1.5px solid var(--border2)" }}>
            <span style={{ fontSize: 12.5, color: "var(--muted)", fontWeight: 600 }}>Current Level:</span>
            <span style={{ fontWeight: 700, color: "var(--gold-dk)" }}>{req.currentLevel}</span>
            {req.ncrPmt < -300 && <span className="badge b-pending">⚠ NCR &lt; -300</span>}
            {req.blocked && <span className="badge b-blocked-order">🔒 Blocked by Admin {req.blockReason ? `— ${req.blockReason}` : ""}</span>}
            {req.validityTo && req.validityTo < TODAY && req.status === "Approved" && <span className="badge b-completed">🏁 Validity Expired</span>}
          </div>

          {/* Sent-back banner for Sales Officer */}
          {soCanResubmit && (
            <div className="alert a-warning" style={{marginBottom:16}}>
              <span style={{fontSize:20}}>↩</span>
              <div style={{flex:1}}>
                <strong>This request was sent back for revision.</strong>
                <div style={{fontSize:12,marginTop:3,fontWeight:500}}>
                  Reason: <em>"{(req.history||[]).filter(h=>h.action==="Sent Back").slice(-1)[0]?.remark || "—"}"</em>
                </div>
                <div style={{fontSize:12,marginTop:2,color:"#92400E"}}>Update the fields below and resubmit to your manager.</div>
              </div>
              <button className="btn btn-warning btn-sm" style={{flexShrink:0,background:"#FEF3C7",color:"#92400E",border:"1.5px solid #FCD34D"}}
                onClick={()=>setShowResubmit(p=>!p)}>
                {showResubmit ? "▲ Hide Edit Form" : "✏️ Edit & Resubmit"}
              </button>
            </div>
          )}

          {/* Inline edit form for resubmission */}
          {soCanResubmit && showResubmit && (
            <div style={{background:"var(--gold-f)",border:"2px solid var(--gold-200)",borderRadius:12,padding:"18px 20px",marginBottom:16}}>
              <div style={{fontFamily:"'Sora',sans-serif",fontWeight:700,fontSize:14,color:"var(--gold-dk)",marginBottom:14}}>
                ✏️ Edit Request — Revision #{(req.history||[]).filter(h=>h.action==="Sent Back").length}
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:12,marginBottom:12}}>
                <div className="fg">
                  <label className="fl fl-req">Order Price (₹)</label>
                  <input className="fc" type="number" value={rsForm.orderPrice}
                    onChange={e=>setRsForm(p=>({...p,orderPrice:+e.target.value}))} />
                </div>
                <div className="fg">
                  <label className="fl fl-req">Quantity (MT)</label>
                  <input className="fc" type="number" value={rsForm.qty}
                    onChange={e=>setRsForm(p=>({...p,qty:+e.target.value}))} />
                </div>
                <div className="fg">
                  <label className="fl">Payment Terms</label>
                  <select className="fc" value={rsForm.paymentTerms}
                    onChange={e=>setRsForm(p=>({...p,paymentTerms:e.target.value}))}>
                    <option>Advance</option><option>Credit</option><option>BG</option><option>LC</option>
                  </select>
                </div>
                <div className="fg">
                  <label className="fl">Validity From</label>
                  <input className="fc" type="date" value={rsForm.validityFrom}
                    onChange={e=>setRsForm(p=>({...p,validityFrom:e.target.value}))} />
                </div>
                <div className="fg">
                  <label className="fl">Validity To</label>
                  <input className="fc" type="date" value={rsForm.validityTo}
                    onChange={e=>setRsForm(p=>({...p,validityTo:e.target.value}))} />
                </div>
              </div>
              <div className="fg" style={{marginBottom:12}}>
                <label className="fl fl-req">Revision Remark <span style={{color:"var(--red)"}}>*</span></label>
                <textarea className="fc" rows={2} value={rsForm.remark}
                  onChange={e=>setRsForm(p=>({...p,remark:e.target.value}))}
                  placeholder="Explain what you changed and why…" />
                {!rsForm.remark.trim() && <span style={{fontSize:11,color:"var(--red)",fontWeight:600}}>Required before resubmitting</span>}
              </div>
              <div style={{display:"flex",gap:10,alignItems:"center"}}>
                <button className="btn btn-primary btn-sm" disabled={!rsForm.remark.trim()} onClick={doResubmit}>
                  🚀 Resubmit to Manager
                </button>
                <button className="btn btn-ghost btn-sm" onClick={()=>setShowResubmit(false)}>Cancel</button>
                <span style={{fontSize:11.5,color:"var(--muted)",marginLeft:4}}>Will be sent back to Area Sales Manager for review</span>
              </div>
            </div>
          )}
          <div className="tabs">{["details", "financials", "history", "attachments"].map(t => <div key={t} className={`tab ${tab === t ? "active" : ""}`} onClick={() => setTab(t)}>{t === "attachments" ? <>📎 Attachments{(req.attachments||[]).length > 0 && <span className="att-count-badge">{(req.attachments||[]).length}</span>}</> : t.charAt(0).toUpperCase() + t.slice(1)}</div>)}</div>
          {tab === "details" && (<><div className="dv-head">📦 Order Information</div><div className="dv-grid"><F label="Customer" value={req.customerName} /><F label="Region" value={req.region} /><F label="Destination" value={req.destination} /><F label="Grade" value={req.grade} /><F label="Material" value={req.materialType} /><F label="Qty (MT)" value={`${req.qty}`} /><F label="Supply From" value={req.supplyFrom} /><F label="Route 1 (Primary Freight)" value={req.route1 ? `${(freightMaster||[]).find(f=>f.id===req.route1)?.route||req.route1} — ${(freightMaster||[]).find(f=>f.id===req.route1)?.routeDesc||""} · ₹${req.primaryFreight}/MT` : req.primaryFreight ? `₹${req.primaryFreight}/MT` : ""} /><F label="Route 2 (Secondary Freight)" value={req.route2 ? `${(freightMaster||[]).find(f=>f.id===req.route2)?.route||req.route2} — ${(freightMaster||[]).find(f=>f.id===req.route2)?.routeDesc||""} · ₹${req.secondaryFreight}/MT` : req.secondaryFreight ? `₹${req.secondaryFreight}/MT` : ""} /><F label="Mode" value={req.mode} /><F label="Source" value={req.source} /><F label="Storage Location" value={req.storageLocation} /><F label="Payment" value={req.paymentTerms} /><F label="TPC Agent" value={req.tpcAgent} /><F label="Validity" value={`${req.validityFrom} → ${req.validityTo}`} /><F label="Unit / Source" value={req.unitSource} /></div></>)}
          {tab === "financials" && (<>
            <div className="dv-head">💰 Pricing</div>
            <div className="dv-grid"><F label="Order Price" value={`₹${req.orderPrice}`} /><F label="Trade Price" value={`₹${req.tradePrice}`} /><F label="Difference" value={`₹${req.difference}`} /><F label="Net of GST" value={`₹${Number(req.netOfGST).toFixed(2)}`} /></div>
            <div className="dv-head">📊 Expenses</div>
            <div className="dv-grid"><F label="Primary Freight" value={`₹${req.primaryFreight}`} /><F label="Secondary Freight" value={`₹${req.secondaryFreight}`} /><F label="CoP" value={`₹${req.costOfProduction}`} /><F label="OP Commission" value={`₹${req.opCommission}`} /><F label="Total Expenses" value={`₹${req.totalExpenses}`} /></div>
            <div className={`ncr-box${req.ncrPmt < -300 ? " critical" : ""}`} style={{ marginTop: 14 }}>
              <div style={{ fontSize: 11.5, fontWeight: 700, textTransform: "uppercase", color: "var(--muted)", marginBottom: 4 }}>NCR Per MT</div>
              <div className={`ncr-val ${req.ncrPmt < 0 ? "ncr-neg" : "ncr-pos"}`}>₹{Number(req.ncrPmt).toFixed(2)}</div>
            </div>
          </>)}
          {tab === "history" && <HistLine history={req.history} users={users} />}
          {tab === "attachments" && (
            <div>
              {atts.length === 0 ? (
                <div className="att-empty">
                  <div className="att-empty-ico">📎</div>
                  <div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 700, fontSize: 15, color: "var(--ink)", marginBottom: 6 }}>No Attachments</div>
                  <div style={{ fontSize: 13, color: "var(--muted)" }}>No documents were attached to this request.</div>
                </div>
              ) : (
                <>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
                    <div style={{ fontSize: 13, fontWeight: 700, color: "var(--ink)" }}>{atts.length} document{atts.length !== 1 ? "s" : ""} attached</div>
                  </div>
                  <div className="att-list">
                    {atts.map(att => (
                      <div key={att.id} className="att-item">
                        <div className="att-icon" style={{ background: getFileColor(att.name) }}>{getFileIcon(att.name)}</div>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div className="att-name">{att.name}</div>
                          <div className="att-meta">{fmtSize(att.size)} · Uploaded {att.uploadedAt}</div>
                        </div>
                        <div className="att-actions">
                          {att.dataUrl && att.dataUrl.startsWith("data:image") && (
                            <button className="btn btn-xs btn-ghost" title="Preview" onClick={() => window.open(att.dataUrl, "_blank")}>👁</button>
                          )}
                          <button className="btn btn-xs btn-ghost" title="Download" onClick={() => downloadAtt(att)}>⬇</button>
                        </div>
                      </div>
                    ))}
                  </div>
                </>
              )}
            </div>
          )}
        </div>
        {/* Footer: manager actions */}
        {canDo && !req.blocked && <div className="m-footer">
          <button className="btn btn-success btn-sm" onClick={() => setAct("approve")}>✓ Approve</button>
          <button className="btn btn-danger btn-sm" onClick={() => setAct("reject")}>✗ Reject</button>
          <button className="btn btn-warning btn-sm" onClick={() => setAct("modify")}>✏ Modify</button>
          <button className="btn btn-secondary btn-sm" onClick={() => setAct("sendback")} style={{ background: "#FEF3C7", color: "#92400E", borderColor: "#FCD34D" }}>↩ Send Back for Revision</button>
          <button className="btn btn-secondary btn-sm" onClick={() => setAct("escalate")}>↑ Escalate</button>
        </div>}
        {/* Footer: SO resubmit — shown when request is sent back */}
        {soCanResubmit && !showResubmit && (
          <div className="m-footer" style={{background:"#FEF3C7",borderTop:"1.5px solid #FCD34D"}}>
            <span style={{fontSize:12.5,color:"#92400E",fontWeight:600,flex:1}}>↩ This request needs your revision before it can proceed.</span>
            <button className="btn btn-warning btn-sm" style={{background:"#D97706",color:"#fff",borderColor:"#D97706"}}
              onClick={()=>{setTab("details");setShowResubmit(true);}}>
              ✏️ Edit & Resubmit
            </button>
          </div>
        )}
      </div>
      {act && (
        <div className="overlay" style={{ zIndex: 600 }}>
          <div className="modal modal-sm">
            <div className="m-hdr"><div style={{ display: "flex", alignItems: "center", gap: 12 }}><div className="m-hdr-accent">{act === "approve" ? "✓" : act === "reject" ? "✗" : act === "sendback" ? "↩" : "✏"}</div><div className="m-title">{{ approve: "Approve", reject: "Reject", modify: "Modify & Change Validity", escalate: "Escalate", sendback: "Send Back for Revision" }[act]}</div></div><button className="m-close" onClick={() => setAct(null)}>×</button></div>
            <div className="m-body">
              {act === "modify" && (<>
                <div className="fg" style={{ marginBottom: 14 }}><label className="fl">Order Price (₹)</label><input className="fc" type="number" value={mp} onChange={e => setMp(+e.target.value)} /></div>
                <div className="fg" style={{ marginBottom: 14 }}><label className="fl">Quantity (MT)</label><input className="fc" type="number" value={mq} onChange={e => setMq(+e.target.value)} /></div>
                <div style={{ background: "var(--gold-f)", border: "1.5px solid var(--gold-200)", borderRadius: 10, padding: "12px 14px", marginBottom: 14 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "var(--gold-dk)", marginBottom: 10 }}>📅 Change Validity Period</div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                    <div className="fg"><label className="fl">Validity From</label><input className="fc" type="date" value={mValidFrom} onChange={e => setMValidFrom(e.target.value)} /></div>
                    <div className="fg"><label className="fl">Validity To</label><input className="fc" type="date" value={mValidTo} onChange={e => setMValidTo(e.target.value)} /></div>
                  </div>
                </div>
              </>)}
              {act === "sendback" && (
                <div className="alert a-warning" style={{ marginBottom: 14 }}>
                  <span>↩</span>
                  <div>This request will be sent back to <strong>{prevLevel(req.currentLevel)}</strong> for revision. A remark is mandatory.</div>
                </div>
              )}
              <div className="fg"><label className="fl">Remark {(act === "reject" || act === "sendback") && <span style={{ color: "var(--red)" }}>*</span>}</label><textarea className="fc" rows={3} value={remark} onChange={e => setRemark(e.target.value)} placeholder={act === "sendback" ? "Specify what needs to be revised..." : "Enter remark..."} /></div>
              {act === "reject" && !remark.trim() && <div className="alert a-danger" style={{ marginTop: 10 }}><span>⚠</span>Remark required.</div>}
              {act === "sendback" && !remark.trim() && <div className="alert a-warning" style={{ marginTop: 10 }}><span>⚠</span>Please specify what needs revision.</div>}
            </div>
            <div className="m-footer"><button className="btn btn-secondary btn-sm" onClick={() => setAct(null)}>Cancel</button><button className={`btn btn-sm ${act === "reject" ? "btn-danger" : act === "approve" ? "btn-success" : act === "sendback" ? "btn-warning" : "btn-warning"}`} onClick={() => doAction(act)} disabled={(act === "reject" || act === "sendback") && !remark.trim()}>{act === "sendback" ? "↩ Send Back" : act === "approve" ? "✓ Confirm Approve" : "Confirm"}</button></div>
          </div>
        </div>
      )}
    </div>
  );
}

// ═══ REQUEST FORM ════════════════════════════════════════
function RequestForm({ user, users, customers, freightMaster, priceMaster, tpcAgents, modeMaster, unitSourceMaster, sourceMaster, storageLocationMaster, opCommissionMaster, onSubmit, onCancel }) {
  const [form, setForm] = useState({ date: new Date().toISOString().split("T")[0], customerCode: "", customerName: "", region: "", zone: "", destination: "", grade: "OPC", materialType: "Bag", qty: "", orderPrice: "", tradePrice: "", difference: "", netOfGST: "", primaryFreight: "", secondaryFreight: "", demrage: "", stationHandling: "", unloadingPrice: "", opCommission: "", spCommission: "", costOfProduction: "", packing: "", totalExpenses: "", ncrPmt: "", paymentTerms: "Advance", supplyFrom: "Aligarh Plant", validityFrom: "", validityTo: "", tpcAgent: "", mode: "", unitSource: "", source: "", storageLocation: "", route1: "", route2: "" });
  const [attachments, setAttachments] = useState([]);
  const [attDrag, setAttDrag] = useState(false);
  const attRef = useRef();

  // Determine which financial fields this user can see
  const currentUserFull = (users||[]).find(u=>u.id===user.id) || user;
  const fieldRights = currentUserFull.formFieldRights || [];
  const canSeeField = (key) => fieldRights.length === 0 || fieldRights.includes(key);

  const FILE_ICONS = { pdf: "📄", doc: "📝", docx: "📝", xls: "📊", xlsx: "📊", png: "🖼️", jpg: "🖼️", jpeg: "🖼️", gif: "🖼️", txt: "📃", zip: "🗜️", rar: "🗜️" };
  const getFileIcon = name => { const ext = (name||"").split(".").pop().toLowerCase(); return FILE_ICONS[ext] || "📎"; };
  const getFileColor = name => { const ext = (name||"").split(".").pop().toLowerCase(); if(["pdf"].includes(ext)) return "#FEF2F2"; if(["xls","xlsx"].includes(ext)) return "#ECFDF5"; if(["png","jpg","jpeg","gif"].includes(ext)) return "#EFF6FF"; if(["doc","docx"].includes(ext)) return "#EFF6FF"; return "var(--cream)"; };
  const fmtSize = b => b < 1024 ? `${b} B` : b < 1048576 ? `${(b/1024).toFixed(1)} KB` : `${(b/1048576).toFixed(1)} MB`;

  const handleAttachFiles = files => {
    const arr = Array.from(files);
    const MAX = 10 * 1024 * 1024; // 10 MB per file
    arr.forEach(f => {
      if (f.size > MAX) { alert(`"${f.name}" exceeds 10 MB limit.`); return; }
      const reader = new FileReader();
      reader.onload = e => {
        setAttachments(prev => [...prev, { id: Date.now() + Math.random(), name: f.name, size: f.size, type: f.type, dataUrl: e.target.result, uploadedAt: new Date().toLocaleString() }]);
      };
      reader.readAsDataURL(f);
    });
  };
  const removeAtt = id => setAttachments(p => p.filter(a => a.id !== id));
  const downloadAtt = att => { const a = document.createElement("a"); a.href = att.dataUrl; a.download = att.name; a.click(); };
  const avCust = customers.filter(c => user.role === "Admin" || c.assignedTo?.includes(user.id));

  // Derive filtered route lists for dropdowns
  const getRouteOptions = (dest, supplyFrom) => {
    return (freightMaster||[]).filter(f => {
      if (!f.to && !f.route) return false;
      if (dest && f.to !== dest) return false;
      return true;
    });
  };

  // Auto-fetch OP Commission from master
  const getOPCommission = (grade, paymentTerms) => {
    const rec = (opCommissionMaster||[]).find(r => r.grade === grade && r.paymentTerms === paymentTerms);
    return rec ? rec.commissionRate : "";
  };

  const update = (field, value) => setForm(prev => {
    const n = { ...prev, [field]: value };
    if (field === "customerCode") { const c = customers.find(c => c.customerCode === value); if (c) { n.customerName = c.customerName; n.region = c.regionDesc; n.zone = c.transportationZone; } }

    // When destination or supplyFrom changes, clear route selections and re-derive freight
    if (field === "destination" || field === "supplyFrom") {
      n.route1 = "";
      n.route2 = "";
      n.primaryFreight = "";
      n.secondaryFreight = "";
    }

    // When route1 is selected, auto-fill primaryFreight
    if (field === "route1") {
      const fr = (freightMaster||[]).find(f => f.id === value || f.route === value);
      n.primaryFreight = fr ? (fr.freightPMT || 0) : "";
    }

    // When route2 is selected, auto-fill secondaryFreight
    if (field === "route2") {
      const fr = (freightMaster||[]).find(f => f.id === value || f.route === value);
      n.secondaryFreight = fr ? (fr.freightPMT || 0) : "";
    }
    // Auto-fill trade price from priceMaster when grade or destination changes
    if (field === "grade" || field === "destination") {
      const dest  = field === "destination" ? value : n.destination;
      const grade = field === "grade"       ? value : n.grade;
      const pr = (priceMaster||[]).find(p =>
        p.grade === grade && (p.salesDistrictName === dest || p.cityName === dest || (dest||"").includes(p.salesDistrictName||"X"))
      );
      if (pr) { n.tradePrice = pr.rate || ""; }
    }
    const g = field === "grade" ? value : n.grade; const m = field === "materialType" ? value : n.materialType;
    if (field === "materialType") { n.packing = PACKING[m] || 0; }
    if (field === "grade" || field === "materialType") { n.packing = PACKING[m] || 0; }
    if (field === "supplyFrom") { n.costOfProduction = COP[field==="supplyFrom"?value:n.supplyFrom] || ""; }

    // Auto-fetch OP Commission when grade or paymentTerms changes
    if (field === "grade" || field === "paymentTerms") {
      const gr = field === "grade" ? value : n.grade;
      const pt = field === "paymentTerms" ? value : n.paymentTerms;
      const opRate = getOPCommission(gr, pt);
      if (opRate !== "") n.opCommission = opRate;
    }

    const op = fmt(field === "orderPrice" ? value : n.orderPrice); const tp = fmt(n.tradePrice);
    n.difference = (op - tp).toFixed(2); n.netOfGST = (op / 1.18).toFixed(2);
    const te = ["primaryFreight", "secondaryFreight", "demrage", "stationHandling", "costOfProduction", "packing", "unloadingPrice", "opCommission", "spCommission"].reduce((s, k) => s + fmt(k === field ? value : n[k]), 0);
    n.totalExpenses = te.toFixed(2); n.ncrPmt = ((op / 1.18) - te / 20).toFixed(2);
    return n;
  });
  const nv = fmt(form.ncrPmt);
  const handleSubmit = () => {
    const id = genId();
    onSubmit({ id, date: form.date, customerCode: form.customerCode, customerName: form.customerName, region: form.region, zone: form.zone, destination: form.destination, grade: form.grade, materialType: form.materialType, qty: fmt(form.qty), orderPrice: fmt(form.orderPrice), tradePrice: fmt(form.tradePrice), difference: fmt(form.difference), netOfGST: fmt(form.netOfGST), primaryFreight: fmt(form.primaryFreight), secondaryFreight: fmt(form.secondaryFreight), demrage: fmt(form.demrage), stationHandling: fmt(form.stationHandling), costOfProduction: fmt(form.costOfProduction), packing: fmt(form.packing), unloadingPrice: fmt(form.unloadingPrice), opCommission: fmt(form.opCommission), spCommission: fmt(form.spCommission), totalExpenses: fmt(form.totalExpenses), ncrPmt: fmt(form.ncrPmt), paymentTerms: form.paymentTerms, supplyFrom: form.supplyFrom, validityFrom: form.validityFrom, validityTo: form.validityTo, tpcAgent: form.tpcAgent, mode: form.mode, unitSource: form.unitSource, source: form.source, storageLocation: form.storageLocation, route1: form.route1, route2: form.route2, attachments: attachments, status: "Pending", currentLevel: "Area Sales Manager", createdBy: user.id, blocked: false, history: [{ level: "Sales Officer", action: "Created", by: user.id, time: new Date().toLocaleString(), remark: "" }] });
  };
  return (
    <div>
      <div className="ph"><div><div className="ph-eyebrow">Workflow</div><div className="ph-title">New Price Request</div></div><div className="ph-actions"><button className="btn btn-secondary" onClick={onCancel}>Cancel</button><button className="btn btn-primary" onClick={handleSubmit} disabled={!form.customerCode || !form.destination || !form.orderPrice}>Submit →</button></div></div>
      <div className="card"><div className="card-body">
        <div className="fsec"><div className="fsec-title">📋 Request Info</div><div className="fgrid2"><div className="fg"><label className="fl">Request ID</label><input className="fc" value={`REQ-2024-${String(reqCtr).padStart(3, "0")}`} disabled /></div><div className="fg"><label className="fl fl-req">Date</label><input className="fc" type="date" value={form.date} onChange={e => update("date", e.target.value)} /></div></div></div>
        <div className="fsec"><div className="fsec-title">👤 Customer</div><div className="fgrid"><div className="fg"><label className="fl fl-req">Customer</label><select className="fc" value={form.customerCode} onChange={e => update("customerCode", e.target.value)}><option value="">-- Select --</option>{avCust.map(c => <option key={c.customerCode} value={c.customerCode}>{c.customerCode} — {c.customerName}</option>)}</select></div><div className="fg"><label className="fl">Name</label><input className="fc" value={form.customerName} readOnly /></div><div className="fg"><label className="fl">Region</label><input className="fc" value={form.region} readOnly /></div><div className="fg"><label className="fl">Zone</label><input className="fc" value={form.zone} readOnly /></div></div></div>
        <div className="fsec"><div className="fsec-title">📦 Order</div><div className="fgrid">

          {/* 1. Destination */}
          <div className="fg"><label className="fl fl-req">Destination</label><select className="fc" value={form.destination} onChange={e => update("destination", e.target.value)}><option value="">-- Select --</option>{[...new Set((freightMaster||[]).map(f=>f.to).filter(Boolean))].sort().map(d => <option key={d} value={d}>{d}</option>)}</select></div>

          {/* 2. Grade */}
          <div className="fg"><label className="fl">Grade</label><select className="fc" value={form.grade} onChange={e => update("grade", e.target.value)}>{"OPC,PPC,ProMaxX".split(",").map(g => <option key={g}>{g}</option>)}</select></div>

          {/* 3. Material */}
          <div className="fg"><label className="fl">Material</label><select className="fc" value={form.materialType} onChange={e => update("materialType", e.target.value)}><option>Bag</option><option>Loose</option></select></div>

          {/* 4. Qty (MT) */}
          <div className="fg"><label className="fl fl-req">Qty (MT)</label><input className="fc" type="number" value={form.qty} onChange={e => update("qty", e.target.value)} /></div>

          {/* 5. Supply From */}
          <div className="fg"><label className="fl">Supply From</label><select className="fc" value={form.supplyFrom} onChange={e => update("supplyFrom", e.target.value)}><option>Depot</option><option>Aligarh Plant</option><option>Morak Plant</option></select></div>

          {/* 6. Route 1 — Primary Freight */}
          <div className="fg" style={{ gridColumn: "span 1" }}>
            <label className="fl" style={{ display: "flex", alignItems: "center", gap: 6 }}>
              Route 1 <span style={{ fontSize: 10, fontWeight: 700, background: "var(--gold-f)", color: "var(--gold-dk)", border: "1px solid var(--gold-200)", borderRadius: 5, padding: "1px 6px", letterSpacing: .3 }}>Primary Freight</span>
            </label>
            <select className="fc" value={form.route1} onChange={e => update("route1", e.target.value)}>
              <option value="">-- Select Route --</option>
              {getRouteOptions(form.destination, form.supplyFrom).map(f => (
                <option key={f.id} value={f.id}>{f.route} — {f.routeDesc || f.to}{f.freightPMT ? ` (₹${f.freightPMT}/MT)` : ""}</option>
              ))}
            </select>
            {form.route1 && <span className="fh" style={{ color: "var(--grn)", fontWeight: 700 }}>₹{form.primaryFreight}/MT auto-filled</span>}
            {form.destination && getRouteOptions(form.destination, form.supplyFrom).length === 0 && <span className="fh" style={{ color: "var(--ora)" }}>No routes found for this destination</span>}
          </div>

          {/* 7. Route 2 — Secondary Freight */}
          <div className="fg" style={{ gridColumn: "span 1" }}>
            <label className="fl" style={{ display: "flex", alignItems: "center", gap: 6 }}>
              Route 2 <span style={{ fontSize: 10, fontWeight: 700, background: "var(--blu-f)", color: "var(--blu)", border: "1px solid #BFDBFE", borderRadius: 5, padding: "1px 6px", letterSpacing: .3 }}>Secondary Freight</span>
            </label>
            <select className="fc" value={form.route2} onChange={e => update("route2", e.target.value)}>
              <option value="">-- Select Route --</option>
              {getRouteOptions(form.destination, form.supplyFrom).map(f => (
                <option key={f.id} value={f.id}>{f.route} — {f.routeDesc || f.to}{f.freightPMT ? ` (₹${f.freightPMT}/MT)` : ""}</option>
              ))}
            </select>
            {form.route2 && <span className="fh" style={{ color: "var(--blu)", fontWeight: 700 }}>₹{form.secondaryFreight}/MT auto-filled</span>}
          </div>

          {/* 8. Mode */}
          <div className="fg"><label className="fl">Mode</label><select className="fc" value={form.mode} onChange={e => update("mode", e.target.value)}><option value="">-- Select --</option>{(modeMaster||[]).map(m => <option key={m.id} value={m.name}>{m.name}</option>)}</select></div>

          {/* 9. Source */}
          <div className="fg"><label className="fl">Source</label><select className="fc" value={form.source} onChange={e => update("source", e.target.value)}><option value="">-- Select --</option>{(sourceMaster||[]).map(s => <option key={s.id} value={s.name}>{s.name}</option>)}</select></div>

          {/* 10. Storage Location */}
          <div className="fg"><label className="fl">Storage Location</label><select className="fc" value={form.storageLocation} onChange={e => update("storageLocation", e.target.value)}><option value="">-- Select --</option>{(storageLocationMaster||[]).map(sl => <option key={sl.id} value={sl.name}>{sl.name}</option>)}</select></div>

          {/* 11. Payment */}
          <div className="fg"><label className="fl">Payment</label><select className="fc" value={form.paymentTerms} onChange={e => update("paymentTerms", e.target.value)}><option>Advance</option><option>Credit</option><option>BG</option><option>LC</option></select></div>

          {/* 12. TPC Agent */}
          <div className="fg"><label className="fl">TPC Agent</label><select className="fc" value={form.tpcAgent} onChange={e => update("tpcAgent", e.target.value)}><option value="">-- None --</option>{tpcAgents.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}</select></div>

          {/* 13. Valid From */}
          <div className="fg"><label className="fl">Valid From</label><input className="fc" type="date" value={form.validityFrom} onChange={e => update("validityFrom", e.target.value)} /></div>

          {/* 14. Valid To */}
          <div className="fg"><label className="fl">Valid To</label><input className="fc" type="date" value={form.validityTo} onChange={e => update("validityTo", e.target.value)} /></div>

          {/* Unit / Source (kept for completeness) */}
          <div className="fg"><label className="fl">Unit / Source</label><select className="fc" value={form.unitSource} onChange={e => update("unitSource", e.target.value)}><option value="">-- Select --</option>{(unitSourceMaster||[]).map(u => <option key={u.id} value={u.name}>{u.name}</option>)}</select></div>

        </div></div>
        <div className="fsec"><div className="fsec-title">💰 Pricing</div>
          {!canSeeField("tradePrice") && !canSeeField("difference") && !canSeeField("netOfGST") && !canSeeField("ncrPmt") && (
            <div className="alert a-info" style={{marginBottom:10}}><span>ℹ️</span>Some pricing fields are restricted for your profile. Contact Admin for full access.</div>
          )}
          <div className="fgrid">
            <div className="fg"><label className="fl fl-req">Order Price (₹)</label><input className="fc" type="number" value={form.orderPrice} onChange={e => update("orderPrice", e.target.value)} /></div>
            {canSeeField("tradePrice") && <div className="fg"><label className="fl">Trade Price (₹)</label><input className="fc" value={form.tradePrice} readOnly /><span className="fh">Auto-filled from Price Master</span></div>}
            {canSeeField("difference") && <div className="fg"><label className="fl">Difference (₹)</label><input className="fc" value={form.difference} readOnly style={{ color: fmt(form.difference) < 0 ? "var(--red)" : "var(--grn)", fontWeight: 700 }} /></div>}
            {canSeeField("netOfGST") && <div className="fg"><label className="fl">Net of GST</label><input className="fc" value={form.netOfGST} readOnly /></div>}
          </div>
        </div>
        <div className="fsec"><div className="fsec-title">📊 Expenses</div><div className="fgrid">
          <div className="fg"><label className="fl">Primary Freight</label><input className="fc" value={form.primaryFreight} readOnly /></div>
          <div className="fg"><label className="fl">Secondary Freight</label><input className="fc" value={form.secondaryFreight} readOnly /></div>
          <div className="fg"><label className="fl">Demrage</label><input className="fc" type="number" value={form.demrage} onChange={e => update("demrage", e.target.value)} /></div>
          <div className="fg"><label className="fl">Station Handling</label><input className="fc" type="number" value={form.stationHandling} onChange={e => update("stationHandling", e.target.value)} /></div>
          {canSeeField("cop") && <div className="fg"><label className="fl">Cost of Production</label><input className="fc" value={form.costOfProduction} readOnly /></div>}
          <div className="fg"><label className="fl">Packing</label><input className="fc" value={form.packing} readOnly /></div>
          <div className="fg"><label className="fl">Unloading</label><input className="fc" type="number" value={form.unloadingPrice} onChange={e => update("unloadingPrice", e.target.value)} /></div>
          {canSeeField("opCommission") && <div className="fg">
            <label className="fl">OP Commission (₹/MT) <span style={{fontSize:10,background:"var(--gold-f)",color:"var(--gold-dk)",borderRadius:4,padding:"1px 5px",fontWeight:700}}>Auto</span></label>
            <input className="fc" type="number" value={form.opCommission} onChange={e => update("opCommission", e.target.value)} />
            <span className="fh">Auto-fetched from OP Commission Master · Grade: {form.grade} · Terms: {form.paymentTerms}</span>
          </div>}
          <div className="fg"><label className="fl">SP Commission</label><input className="fc" type="number" value={form.spCommission} onChange={e => update("spCommission", e.target.value)} /></div>
        </div></div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
          <div className="card" style={{ border: "1.5px solid var(--border2)" }}><div className="card-body"><div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: .5, color: "var(--muted)", marginBottom: 6 }}>Total Expenses</div><div style={{ fontFamily: "'Sora',sans-serif", fontSize: 28, fontWeight: 800 }}>₹{form.totalExpenses}</div></div></div>
          {canSeeField("ncrPmt") && <div className={`ncr-box${nv < -300 ? " critical" : ""}`}><div style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: .5, color: "var(--muted)", marginBottom: 6 }}>NCR Per MT</div><div className={`ncr-val ${nv < 0 ? "ncr-neg" : "ncr-pos"}`}>₹{form.ncrPmt}</div>{nv < -300 && <div style={{ fontSize: 12, color: "var(--red)", marginTop: 6, fontWeight: 600 }}>⚠ Below -300 threshold</div>}</div>}
        </div>
        <div className="fsec" style={{ marginTop: 16 }}>
          <div className="fsec-title">📎 Attachments {attachments.length > 0 && <span className="att-count-badge">{attachments.length}</span>}</div>
          <input ref={attRef} type="file" multiple style={{ display: "none" }} onChange={e => { handleAttachFiles(e.target.files); e.target.value = ""; }} />
          <div
            className={`att-zone${attDrag ? " drag" : ""}`}
            onClick={() => attRef.current.click()}
            onDragOver={e => { e.preventDefault(); setAttDrag(true); }}
            onDragLeave={() => setAttDrag(false)}
            onDrop={e => { e.preventDefault(); setAttDrag(false); handleAttachFiles(e.dataTransfer.files); }}
          >
            <div className="att-zone-ico">📂</div>
            <div className="att-zone-title">Drop files here or click to browse</div>
            <div className="att-zone-sub">PDF, Word, Excel, Images · Max 10 MB per file</div>
          </div>
          {attachments.length > 0 && (
            <div className="att-list">
              {attachments.map(att => (
                <div key={att.id} className="att-item">
                  <div className="att-icon" style={{ background: getFileColor(att.name) }}>{getFileIcon(att.name)}</div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div className="att-name">{att.name}</div>
                    <div className="att-meta">{fmtSize(att.size)} · {att.uploadedAt}</div>
                  </div>
                  <div className="att-actions">
                    <button className="btn btn-xs btn-ghost" title="Download" onClick={e => { e.stopPropagation(); downloadAtt(att); }}>⬇</button>
                    <button className="btn btn-xs btn-ghost" title="Remove" style={{ color: "var(--red)" }} onClick={e => { e.stopPropagation(); removeAtt(att.id); }}>🗑</button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div></div>
    </div>
  );
}

// ═══ CUSTOMER MASTER ════════════════════════════════════
function CustFormModal({ customer, users, onSave, onClose }) {
  const isEdit = !!(customer?.customerCode);
  const [form, setForm] = useState(customer ? { ...customer } : { ...EMPTY_CUST, id: Date.now() });
  const [sec, setSec] = useState(0);
  const update = (k, v) => setForm(p => ({ ...p, [k]: v }));
  const toggleUser = uid => setForm(p => ({ ...p, assignedTo: p.assignedTo?.includes(uid) ? p.assignedTo.filter(x => x !== uid) : [...(p.assignedTo || []), uid] }));
  const s = CUST_FORM_SECTIONS[sec];
  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal modal-xl" style={{ maxHeight: "94vh" }}>
        <div className="m-hdr"><div style={{ display: "flex", alignItems: "center", gap: 12 }}><div className="m-hdr-accent">{isEdit ? "✏️" : "➕"}</div><div><div className="m-title">{isEdit ? "Edit" : "New"} Customer</div><div className="m-sub">{isEdit ? `${form.customerCode} — ${form.customerName}` : "Fill required fields"}</div></div></div><button className="m-close" onClick={onClose}>×</button></div>
        <div style={{ display: "flex", flex: 1, overflow: "hidden", minHeight: 0 }}>
          <div className="cm-sect-nav">{CUST_FORM_SECTIONS.map((s2, i) => <div key={i} className={`cm-sn-item ${i === sec ? "active" : ""}`} onClick={() => setSec(i)}><span style={{ fontSize: 16 }}>{s2.icon}</span><span>{s2.title}</span></div>)}</div>
          <div className="m-body" style={{ flex: 1 }}>
            <div className="fsec"><div className="fsec-title">{s.icon} {s.title}</div>
              {sec < CUST_FORM_SECTIONS.length - 1 ? (<div className="fgrid">{s.fields.map(f => <div className="fg" key={f.key}><label className={`fl${f.req ? " fl-req" : ""}`}>{f.label}</label>{f.type === "select" ? <select className="fc" value={form[f.key] || ""} onChange={e => update(f.key, e.target.value)}><option value="">-- Select --</option>{f.options.map(o => <option key={o}>{o}</option>)}</select> : <input className="fc" type={f.type || "text"} value={form[f.key] || ""} onChange={e => update(f.key, e.target.value)} />}</div>)}</div>)
                : (<div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(240px,1fr))", gap: 10 }}>{users.map(u => <label key={u.id} style={{ display: "flex", alignItems: "center", gap: 10, cursor: "pointer", padding: "10px 14px", borderRadius: 10, border: "1.5px solid", borderColor: form.assignedTo?.includes(u.id) ? "var(--gold-200)" : "var(--border2)", background: form.assignedTo?.includes(u.id) ? "var(--gold-f)" : "var(--white)", transition: "all .15s" }}><input type="checkbox" checked={!!form.assignedTo?.includes(u.id)} onChange={() => toggleUser(u.id)} style={{ accentColor: "var(--gold)", width: 16, height: 16 }} /><div><div style={{ fontWeight: 700, fontSize: 13 }}>{u.name}</div><div style={{ fontSize: 11, color: "var(--muted)" }}>{u.role}</div></div></label>)}</div>)}
            </div>
          </div>
        </div>
        <div className="m-footer"><button className="btn btn-secondary btn-sm" onClick={() => setSec(Math.max(0, sec - 1))} disabled={sec === 0}>← Prev</button><button className="btn btn-secondary btn-sm" onClick={() => setSec(Math.min(CUST_FORM_SECTIONS.length - 1, sec + 1))} disabled={sec === CUST_FORM_SECTIONS.length - 1}>Next →</button><span style={{ fontSize: 11, color: "var(--muted)", margin: "0 8px", alignSelf: "center" }}>{sec + 1}/{CUST_FORM_SECTIONS.length}</span><div style={{ flex: 1 }} /><button className="btn btn-secondary btn-sm" onClick={onClose}>Cancel</button><button className="btn btn-primary btn-sm" onClick={() => { if (!form.customerCode || !form.customerName) return; onSave(form); onClose(); }} disabled={!form.customerCode || !form.customerName}>{isEdit ? "💾 Update" : "✅ Save"}</button></div>
      </div>
    </div>
  );
}

function CustomerMasterPage({ customers, setCustomers, users }) {
  const [search, setSearch] = useState(""); const [fStatus, setFStatus] = useState(""); const [fRegion, setFRegion] = useState(""); const [fPlant, setFPlant] = useState("");
  const [pg, setPg] = useState(1); const PS = 10;
  const [modal, setModal] = useState(null); const [visCols, setVisCols] = useState(DEFAULT_CM_COLS); const [selRows, setSelRows] = useState(new Set());
  const { exportCsv, exportTemplate, ExportModal, ToastEl } = useExport();
  const [toast2, showToast2] = useToast();

  const regions = [...new Set(customers.map(c => c.regionDesc).filter(Boolean))];
  const plants = [...new Set(customers.map(c => c.plant).filter(Boolean))];
  const filtered = customers.filter(c => {
    if (fStatus && c.status !== fStatus) return false; if (fRegion && c.regionDesc !== fRegion) return false; if (fPlant && c.plant !== fPlant) return false;
    if (search) { const q = search.toLowerCase(); return c.customerName?.toLowerCase().includes(q) || c.customerCode?.toLowerCase().includes(q) || c.mobileNo?.includes(q) || c.gstin?.toLowerCase().includes(q) || c.city?.toLowerCase().includes(q); }
    return true;
  });
  const totalPages = Math.ceil(filtered.length / PS); const paginated = filtered.slice((pg - 1) * PS, pg * PS);
  const colDefs = ALL_CM_COLS.filter(c => visCols.includes(c.key));
  const handleSave = d => { setCustomers(prev => { const i = prev.findIndex(c => c.id === d.id); if (i >= 0) { const n = [...prev]; n[i] = d; return n; } return [d, ...prev]; }); showToast2(d.customerName + " saved!"); setModal(null); };
  const handleDelete = id => { setCustomers(prev => prev.filter(c => c.id !== id)); showToast2("Deleted.", "danger"); };
  const handleImport = recs => { setCustomers(prev => { const map = new Map(prev.map(c => [c.customerCode, c])); recs.forEach(r => { if (r.customerCode) map.set(r.customerCode, { ...EMPTY_CUST, ...r, id: map.get(r.customerCode)?.id || r.id, assignedTo: [] }); }); return [...map.values()]; }); showToast2(`Imported!`); };
  const exportData = selRows.size > 0 ? customers.filter(c => selRows.has(c.id)) : filtered;

  return (
    <div>
      {toast2 && <Toast msg={toast2.msg} type={toast2.type} />}
      {ToastEl}
      <div className="ph">
        <div><div className="ph-eyebrow">Master Data</div><div className="ph-title">Customer Master</div><div className="ph-sub">{filtered.length} customers{selRows.size > 0 && ` · ${selRows.size} selected`}</div></div>
        <div className="ph-actions">
          <button className="btn btn-secondary btn-sm" onClick={() => exportTemplate(ALL_CM_COLS, "customer_master_template.csv")}>📋 Template</button>
          <button className="btn btn-secondary btn-sm" onClick={() => setModal({ type: "import" })}>📤 Import</button>
          <button className="btn btn-secondary btn-sm" onClick={() => exportCsv(exportData, ALL_CM_COLS, `customers_${new Date().toISOString().split("T")[0]}.csv`, `Customers (${exportData.length})`)}>⬇ Export{selRows.size > 0 ? ` (${selRows.size})` : ""}</button>
          <button className="btn btn-primary btn-sm" onClick={() => setModal({ type: "form", data: null })}>+ Add</button>
        </div>
      </div>
      <div className="kpi-grid" style={{ gridTemplateColumns: "repeat(4,1fr)" }}>
        {[{ l: "Total", v: customers.length, c: "var(--gold)", bg: "var(--gold-f)", i: "👥" }, { l: "Active", v: customers.filter(c => c.status === "Active").length, c: "var(--grn)", bg: "var(--grn-f)", i: "✅" }, { l: "Inactive", v: customers.filter(c => c.status === "Inactive").length, c: "var(--red)", bg: "var(--red-f)", i: "❌" }, { l: "Blocked", v: customers.filter(c => c.status === "Blocked").length, c: "#F59E0B", bg: "#FFF8E1", i: "🔒" }].map((k, i) => (
          <div key={i} className="kpi"><div className="kpi-accent" style={{ background: k.c }} /><div className="kpi-icon" style={{ background: k.bg }}>{k.i}</div><div className="kpi-val">{k.v}</div><div className="kpi-label">{k.l}</div></div>
        ))}
      </div>
      <div className="card" style={{ marginBottom: 14 }}><div className="card-body" style={{ padding: "14px 18px" }}>
        <div className="sf">
          <div className="si-wrap"><span className="si-ico">🔍</span><input className="si" placeholder="Search name, code, mobile, GSTIN…" value={search} onChange={e => { setSearch(e.target.value); setPg(1); }} /></div>
          <select className="fsel" value={fStatus} onChange={e => { setFStatus(e.target.value); setPg(1); }}><option value="">All Status</option><option>Active</option><option>Inactive</option><option>Blocked</option></select>
          <select className="fsel" value={fRegion} onChange={e => { setFRegion(e.target.value); setPg(1); }}><option value="">All Regions</option>{regions.map(r => <option key={r}>{r}</option>)}</select>
          <select className="fsel" value={fPlant} onChange={e => { setFPlant(e.target.value); setPg(1); }}><option value="">All Plants</option>{plants.map(p => <option key={p}>{p}</option>)}</select>
          {(search || fStatus || fRegion || fPlant) && <button className="btn btn-ghost btn-sm" onClick={() => { setSearch(""); setFStatus(""); setFRegion(""); setFPlant(""); }}>✕ Clear</button>}
        </div>
      </div></div>
      <div className="card">
        <div className="tw"><table>
          <thead><tr>
            <th style={{ width: 36 }}><input type="checkbox" checked={paginated.length > 0 && selRows.size === paginated.length} onChange={() => setSelRows(p => p.size === paginated.length ? new Set() : new Set(paginated.map(c => c.id)))} style={{ accentColor: "var(--gold)" }} /></th>
            {colDefs.map(c => <th key={c.key}>{c.label}</th>)}
            <th>Actions</th>
          </tr></thead>
          <tbody>
            {paginated.length === 0 && <tr><td colSpan={colDefs.length + 2} style={{ textAlign: "center", padding: "48px", color: "var(--muted)" }}><div style={{ fontSize: 36, marginBottom: 10 }}>👥</div><div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 700, fontSize: 15 }}>No customers found</div></td></tr>}
            {paginated.map(c => (
              <tr key={c.id}>
                <td><input type="checkbox" checked={selRows.has(c.id)} onChange={() => setSelRows(p => { const n = new Set(p); n.has(c.id) ? n.delete(c.id) : n.add(c.id); return n; })} style={{ accentColor: "var(--gold)" }} /></td>
                {colDefs.map(col => (
                  <td key={col.key} style={{ cursor: "pointer" }} onClick={() => setModal({ type: "form", data: c })}>
                    {col.key === "customerCode" && <span style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: "var(--gold-dk)", fontSize: 13 }}>{c[col.key]}</span>}
                    {col.key === "customerName" && <div><div style={{ fontWeight: 700, fontSize: 13 }}>{c[col.key]}</div><div style={{ fontSize: 11, color: "var(--muted)" }}>{c.proprietorDirector}</div></div>}
                    {col.key === "status" && <span className={`badge ${c.status === "Active" ? "b-active" : c.status === "Blocked" ? "b-blocked" : "b-inactive"}`}>{c[col.key]}</span>}
                    {(col.key === "creditLimit" || col.key === "securityDeposit") && c[col.key] && <span style={{ fontWeight: 700 }}>₹{Number(c[col.key]).toLocaleString()}</span>}
                    {!["customerCode", "customerName", "status", "creditLimit", "securityDeposit"].includes(col.key) && <span style={{ fontSize: 12.5, fontWeight: 500 }}>{c[col.key] || <span style={{ color: "var(--border)" }}>—</span>}</span>}
                  </td>
                ))}
                <td><div style={{ display: "flex", gap: 4 }}><button className="btn btn-xs btn-ghost" onClick={() => setModal({ type: "form", data: c })}>✏️</button><button className="btn btn-xs btn-ghost" style={{ color: "var(--red)" }} onClick={() => setModal({ type: "delete", data: c })}>🗑</button></div></td>
              </tr>
            ))}
          </tbody>
        </table></div>
        <div style={{ padding: "0 20px 16px" }}>
          <div className="pag">
            <div className="pag-info">Showing {Math.min((pg - 1) * PS + 1, filtered.length)}–{Math.min(pg * PS, filtered.length)} of {filtered.length}</div>
            <div className="pag-btns"><button className="pb" disabled={pg === 1} onClick={() => setPg(1)}>«</button><button className="pb" disabled={pg === 1} onClick={() => setPg(p => p - 1)}>‹</button>{Array.from({ length: Math.min(totalPages, 5) }, (_, i) => { const p2 = Math.max(1, Math.min(pg - 2, totalPages - 4)) + i; return p2 >= 1 && p2 <= totalPages ? <button key={p2} className={`pb ${p2 === pg ? "active" : ""}`} onClick={() => setPg(p2)}>{p2}</button> : null; })}<button className="pb" disabled={pg === totalPages || totalPages === 0} onClick={() => setPg(p => p + 1)}>›</button><button className="pb" disabled={pg === totalPages || totalPages === 0} onClick={() => setPg(totalPages)}>»</button></div>
          </div>
        </div>
      </div>
      {modal?.type === "form" && <CustFormModal customer={modal.data} users={users} onSave={handleSave} onClose={() => setModal(null)} />}
      {modal?.type === "import" && <ImportModal title="Customer Master" colDefs={ALL_CM_COLS} previewCols={["customerCode", "customerName", "regionDesc", "mobileNo", "status"]} onImport={handleImport} onClose={() => setModal(null)} />}
      {modal?.type === "delete" && <ConfirmDelete label={`${modal.data.customerName} (${modal.data.customerCode})`} onConfirm={() => { handleDelete(modal.data.id); setModal(null); }} onClose={() => setModal(null)} />}
      {ExportModal}
    </div>
  );
}

// ═══ ADMIN MASTERS ══════════════════════════════════════
function InlineEditTable({ title, icon, data, setData, cols, keyField, numFields = [], importTitle, previewCols }) {
  const [editRow, setEditRow] = useState(null); const [delRow, setDelRow] = useState(null); const [importOpen, setImportOpen] = useState(false); const [search, setSearch] = useState("");
  const { exportCsv, exportTemplate, ExportModal, ToastEl } = useExport();
  const [toast2, showToast2] = useToast();
  const EMPTY = Object.fromEntries(cols.map(c => [c.key, ""]));
  const saveRow = () => { const d = editRow.data; if (!d[keyField]) return; const proc = { ...d }; numFields.forEach(f => { proc[f] = Number(d[f]) || 0; }); if (editRow.idx === -1) { setData(p => [...p, proc]); showToast2("Added!"); } else { setData(p => p.map((r, i) => i === editRow.idx ? proc : r)); showToast2("Updated!"); } setEditRow(null); };
  const delFn = idx => { setData(p => p.filter((_, i) => i !== idx)); showToast2("Deleted.", "danger"); };
  const handleImport = recs => { setData(prev => { const map = new Map(prev.map(r => [r[keyField], r])); recs.forEach(r => { if (r[keyField]) { const proc = { ...r }; numFields.forEach(f => { proc[f] = Number(r[f]) || 0; }); map.set(r[keyField], proc); } }); return [...map.values()]; }); showToast2("Imported!"); };
  const filtered = data.filter(r => !search || cols.some(c => String(r[c.key] || "").toLowerCase().includes(search.toLowerCase())));
  return (
    <div>
      {toast2 && <Toast msg={toast2.msg} type={toast2.type} />}
      {ToastEl}
      <div className="card">
        <div className="card-hdr">
          <div><div className="card-title">{icon} {title}</div><div className="card-sub">{data.length} records</div></div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
            <div className="si-wrap" style={{ minWidth: 160 }}><span className="si-ico">🔍</span><input className="si" placeholder="Search…" value={search} onChange={e => setSearch(e.target.value)} /></div>
            <button className="btn btn-secondary btn-sm" onClick={() => exportTemplate(cols, `${title.toLowerCase().replace(/\s+/g, "_")}_template.csv`)}>📋 Template</button>
            <button className="btn btn-secondary btn-sm" onClick={() => setImportOpen(true)}>📤 Import</button>
            <button className="btn btn-secondary btn-sm" onClick={() => exportCsv(data, cols, `${title.toLowerCase().replace(/\s+/g, "_")}_${new Date().toISOString().split("T")[0]}.csv`, title)}>⬇ Export</button>
            <button className="btn btn-primary btn-sm" onClick={() => setEditRow({ idx: -1, data: { ...EMPTY } })}>+ Add</button>
          </div>
        </div>
        <div className="tw"><table>
          <thead><tr>{cols.map(c => <th key={c.key}>{c.label}</th>)}<th>Actions</th></tr></thead>
          <tbody>
            {editRow?.idx === -1 && (<tr style={{ background: "var(--gold-f)" }}>{cols.map(c => <td key={c.key}><input className="fc" type={numFields.includes(c.key) ? "number" : "text"} value={editRow.data[c.key]} onChange={e => setEditRow(r => ({ ...r, data: { ...r.data, [c.key]: e.target.value } }))} placeholder={c.label} style={{ padding: "6px 10px", fontSize: 12 }} /></td>)}<td><div style={{ display: "flex", gap: 4 }}><button className="btn btn-xs btn-success" onClick={saveRow}>✓</button><button className="btn btn-xs btn-ghost" onClick={() => setEditRow(null)}>✕</button></div></td></tr>)}
            {filtered.length === 0 && !editRow && <tr><td colSpan={cols.length + 1} style={{ textAlign: "center", padding: "32px", color: "var(--muted)" }}>No records</td></tr>}
            {filtered.map((r, i) => { const ri = data.indexOf(r); const ie = editRow?.idx === ri; return (<tr key={ri} style={ie ? { background: "var(--gold-f)" } : {}}>{cols.map(c => <td key={c.key}>{ie ? <input className="fc" type={numFields.includes(c.key) ? "number" : "text"} value={editRow.data[c.key]} onChange={e => setEditRow(r => ({ ...r, data: { ...r.data, [c.key]: e.target.value } }))} style={{ padding: "6px 10px", fontSize: 12 }} /> : <span style={{ fontWeight: c.key === keyField ? 800 : 500, color: c.key === keyField ? "var(--gold-dk)" : "var(--ink)", fontSize: 13 }}>{numFields.includes(c.key) ? `₹${r[c.key]}` : r[c.key] || "—"}</span>}</td>)}<td>{ie ? <div style={{ display: "flex", gap: 4 }}><button className="btn btn-xs btn-success" onClick={saveRow}>✓</button><button className="btn btn-xs btn-ghost" onClick={() => setEditRow(null)}>✕</button></div> : <div style={{ display: "flex", gap: 4 }}><button className="btn btn-xs btn-ghost" onClick={() => setEditRow({ idx: ri, data: { ...r } })}>✏️</button><button className="btn btn-xs btn-ghost" style={{ color: "var(--red)" }} onClick={() => setDelRow({ idx: ri, label: r[cols[1]?.key] || ri })}>🗑</button></div>}</td></tr>); })}
          </tbody>
        </table></div>
      </div>
      {importOpen && <ImportModal title={importTitle || title} colDefs={cols} previewCols={previewCols || cols.map(c => c.key)} onImport={handleImport} onClose={() => setImportOpen(false)} />}
      {delRow && <ConfirmDelete label={delRow.label} onConfirm={() => delFn(delRow.idx)} onClose={() => setDelRow(null)} />}
      {ExportModal}
    </div>
  );
}

function TradePriceMasterPanel({ tradePrices, setTradePrices }) {
  const toRows = tp => Object.entries(tp).flatMap(([g, types]) => Object.entries(types).map(([m, p]) => ({ grade: g, materialType: m, price: p })));
  const toObj = rows => { const o = {}; rows.forEach(r => { if (!o[r.grade]) o[r.grade] = {}; o[r.grade][r.materialType] = Number(r.price) || 0; }); return o; };
  const [rows, setRows] = useState(toRows(tradePrices));
  const setRowsAndSync = fn => { const r = typeof fn === "function" ? fn(rows) : fn; setRows(r); setTradePrices(toObj(r)); };
  return <InlineEditTable title="Trade Price Master" icon="💰" data={rows} setData={setRowsAndSync} cols={TRADE_PRICE_COLS} keyField="grade" numFields={["price"]} previewCols={["grade", "materialType", "price"]} />;
}

function PwModal({ user: u, onSave, onClose }) {
  const [pw, setPw] = useState(""); const [pw2, setPw2] = useState(""); const [show, setShow] = useState(false);
  const valid = pw.length >= 6 && pw === pw2;
  return (
    <div className="overlay" style={{ zIndex: 700 }} onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal modal-xs">
        <div className="m-hdr"><div style={{ display: "flex", alignItems: "center", gap: 12 }}><div className="m-hdr-accent">🔑</div><div><div className="m-title">Change Password</div><div className="m-sub">{u.name}</div></div></div><button className="m-close" onClick={onClose}>×</button></div>
        <div className="m-body">
          <div className="fg" style={{ marginBottom: 14 }}><label className="fl fl-req">New Password</label><div className="pass-wrap"><input className="fc" type={show ? "text" : "password"} value={pw} onChange={e => setPw(e.target.value)} placeholder="Min 6 characters" style={{ paddingRight: 44 }} /><button className="pass-toggle" onClick={() => setShow(p => !p)}>{show ? "🙈" : "👁"}</button></div></div>
          <div className="fg"><label className="fl fl-req">Confirm</label><input className="fc" type={show ? "text" : "password"} value={pw2} onChange={e => setPw2(e.target.value)} /></div>
          {pw && pw2 && !valid && <div className="alert a-danger" style={{ marginTop: 10 }}><span>⚠</span>{pw.length < 6 ? "Min 6 chars." : "Passwords do not match."}</div>}
        </div>
        <div className="m-footer"><button className="btn btn-secondary btn-sm" onClick={onClose}>Cancel</button><button className="btn btn-primary btn-sm" disabled={!valid} onClick={() => { onSave(pw); onClose(); }}>✅ Save</button></div>
      </div>
    </div>
  );
}

function UserFormModal({ user: editUser, allUsers, customers, locationMaster, onSave, onClose }) {
  const isEdit = !!editUser?.id;
  const emptyForm = { id:"", name:"", role:"Sales Officer", email:"", password:"", status:"Active", assignedCustomers:[], designation:"",
    allowedRegions:[], allowedClusters:[], allowedSalesOffices:[], allowedDistricts:[], allowedTehsils:[], allowedCities:[],
    allowedPages:[], formFieldRights:[], analyticsRights:[] }; // empty allowedPages = access ALL pages; empty formFieldRights = see ALL fields
  const [form, setForm] = useState(editUser ? { ...emptyForm, ...editUser } : emptyForm);
  const [errors, setErrors] = useState({}); const [showPw, setShowPw] = useState(false);
  const [locTab, setLocTab] = useState("salesOffices");
  const lm = locationMaster || {};

  const validate = () => { const e = {}; if (!form.id.trim()) e.id="Required"; if (!form.name.trim()) e.name="Required"; if (!form.email.trim()) e.email="Required"; if (!isEdit && !form.password.trim()) e.password="Required"; if (!isEdit && form.password.length<6) e.password="Min 6 chars"; if (!isEdit && allUsers.find(u=>u.id===form.id.trim())) e.id="ID exists"; setErrors(e); return !Object.keys(e).length; };
  const toggleCust = code => setForm(p=>({ ...p, assignedCustomers: p.assignedCustomers.includes(code)?p.assignedCustomers.filter(c=>c!==code):[...p.assignedCustomers,code] }));
  const toggleLoc = (field, id) => setForm(p=>({ ...p, [field]: (p[field]||[]).includes(id)?(p[field]||[]).filter(x=>x!==id):[...(p[field]||[]),id] }));
  const clearLoc = (field) => setForm(p=>({ ...p, [field]:[] }));
  const selectAllLoc = (field, ids) => setForm(p=>({ ...p, [field]:ids }));

  const LOC_TABS = [
    {key:"allowedRegions",    label:"Region",       items: lm.regions||[]},
    {key:"allowedClusters",   label:"Cluster/Zone", items: lm.clusters||[]},
    {key:"allowedSalesOffices",label:"Sales Office", items: lm.salesOffices||[]},
    {key:"allowedDistricts",  label:"District",     items: lm.districts||[]},
    {key:"allowedTehsils",    label:"Tehsil",       items: lm.tehsils||[]},
    {key:"allowedCities",     label:"City",         items: lm.cities||[]},
  ];
  const curTab  = LOC_TABS.find(t=>t.key===locTab)||LOC_TABS[0];
  const selIds  = form[curTab.key]||[];
  const allIds  = curTab.items.map(i=>i.id);
  const isAdminRole = form.role==="Admin"||form.role==="Zonal Head"||form.role==="Sales & Accounts";

  const ChkBox = ({item, field}) => {
    const sel = (form[field]||[]).includes(item.id);
    return (
      <label style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer",padding:"9px 12px",borderRadius:9,
        background:sel?"var(--gold-f)":"var(--white)",border:"1.5px solid",
        borderColor:sel?"var(--gold-200)":"var(--border2)",transition:"all .15s",opacity:isAdminRole?.5:1}}>
        <input type="checkbox" checked={sel} disabled={isAdminRole} onChange={()=>toggleLoc(field,item.id)} style={{accentColor:"var(--gold)",width:16,height:16}}/>
        <div style={{fontWeight:700,fontSize:13}}>{item.name}</div>
      </label>
    );
  };

  return (
    <div className="overlay" onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div className="modal" style={{maxWidth:760,maxHeight:"92vh",display:"flex",flexDirection:"column"}}>
        <div className="m-hdr"><div style={{display:"flex",alignItems:"center",gap:12}}><div className="m-hdr-accent">{isEdit?"✏️":"➕"}</div><div><div className="m-title">{isEdit?"Edit":"Add"} User</div></div></div><button className="m-close" onClick={onClose}>×</button></div>
        <div className="m-body" style={{overflowY:"auto",flex:1}}>

          {/* User Info */}
          <div className="fsec"><div className="fsec-title">👤 User Info</div>
            <div className="fgrid">
              <div className="fg"><label className="fl fl-req">User ID</label><input className="fc" value={form.id} onChange={e=>setForm(p=>({...p,id:e.target.value.toUpperCase()}))} disabled={isEdit} style={isEdit?{background:"var(--cream)",color:"var(--muted)"}:{}}/>{errors.id&&<span style={{fontSize:11,color:"var(--red)",fontWeight:600}}>{errors.id}</span>}</div>
              <div className="fg"><label className="fl fl-req">Full Name</label><input className="fc" value={form.name} onChange={e=>setForm(p=>({...p,name:e.target.value}))}/>{errors.name&&<span style={{fontSize:11,color:"var(--red)"}}>{errors.name}</span>}</div>
              <div className="fg"><label className="fl fl-req">Email</label><input className="fc" type="email" value={form.email} onChange={e=>setForm(p=>({...p,email:e.target.value}))}/></div>
              <div className="fg"><label className="fl fl-req">Role</label><select className="fc" value={form.role} onChange={e=>setForm(p=>({...p,role:e.target.value}))}>{ROLES.map(r=><option key={r}>{r}</option>)}</select></div>
              <div className="fg"><label className="fl">Status</label><select className="fc" value={form.status||"Active"} onChange={e=>setForm(p=>({...p,status:e.target.value}))}><option>Active</option><option>Inactive</option></select></div>
              <div className="fg"><label className="fl">Designation</label><input className="fc" value={form.designation||""} onChange={e=>setForm(p=>({...p,designation:e.target.value}))}/></div>
              {!isEdit&&(<div className="fg"><label className="fl fl-req">Password</label><div className="pass-wrap"><input className="fc" type={showPw?"text":"password"} value={form.password} onChange={e=>setForm(p=>({...p,password:e.target.value}))} placeholder="Min 6 chars" style={{paddingRight:44}}/><button className="pass-toggle" onClick={()=>setShowPw(p=>!p)}>{showPw?"🙈":"👁"}</button></div>{errors.password&&<span style={{fontSize:11,color:"var(--red)"}}>{errors.password}</span>}</div>)}
            </div>
          </div>

          {/* Location Access Rights */}
          <div className="fsec">
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
              <div className="fsec-title" style={{marginBottom:0}}>📍 Location Access Rights</div>
              {isAdminRole
                ? <span style={{fontSize:12,color:"var(--grn)",fontWeight:600}}>✅ All locations (role has full access)</span>
                : <span style={{fontSize:11,color:"var(--muted)"}}>Empty = access ALL · Select specific items to restrict</span>}
            </div>
            {/* Level tabs */}
            <div style={{display:"flex",gap:6,marginBottom:12,flexWrap:"wrap"}}>
              {LOC_TABS.map(t=>{
                const cnt=(form[t.key]||[]).length;
                return(
                  <button key={t.key} onClick={()=>setLocTab(t.key)}
                    style={{padding:"5px 14px",borderRadius:8,border:`1.5px solid ${locTab===t.key?"var(--gold)":"var(--border2)"}`,
                      background:locTab===t.key?"var(--gold-f)":"var(--white)",fontWeight:700,fontSize:12,
                      color:locTab===t.key?"var(--gold-dk)":"var(--ink)",cursor:"pointer",position:"relative"}}>
                    {t.label}
                    {cnt>0&&<span style={{marginLeft:5,background:"var(--mangRed,#C0142E)",color:"#fff",borderRadius:10,padding:"1px 7px",fontSize:10,fontWeight:800}}>{cnt}</span>}
                  </button>
                );
              })}
            </div>
            {/* Items grid */}
            <div style={{border:"1.5px solid var(--border2)",borderRadius:10,overflow:"hidden"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 14px",background:"var(--cream)",borderBottom:"1px solid var(--border2)"}}>
                <span style={{fontWeight:700,fontSize:13}}>{curTab.label} ({curTab.items.length} total · {selIds.length} selected)</span>
                {!isAdminRole&&(
                  <div style={{display:"flex",gap:8}}>
                    <button onClick={()=>selectAllLoc(curTab.key,allIds)} style={{fontSize:11,fontWeight:700,color:"var(--gold-dk)",background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>Select All</button>
                    <button onClick={()=>clearLoc(curTab.key)} style={{fontSize:11,fontWeight:700,color:"var(--red)",background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>Clear</button>
                  </div>
                )}
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:8,padding:10,maxHeight:160,overflowY:"auto",background:"var(--white)"}}>
                {curTab.items.length===0&&<div style={{gridColumn:"1/-1",textAlign:"center",padding:20,color:"var(--muted)",fontSize:13}}>No {curTab.label} entries. Add them in Location Master first.</div>}
                {curTab.items.map(item=><ChkBox key={item.id} item={item} field={curTab.key}/>)}
              </div>
            </div>
          </div>

          {/* Page Access Rights */}
          <div className="fsec">
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
              <div className="fsec-title" style={{marginBottom:0}}>🔐 Page Access Rights</div>
              {isAdminRole
                ? <span style={{fontSize:12,color:"var(--grn)",fontWeight:600}}>✅ All pages (role has full access)</span>
                : <span style={{fontSize:11,color:"var(--muted)"}}>Empty = access ALL · Uncheck to restrict</span>}
            </div>
            {/* Group pages */}
            {["Main","Analytics","Admin"].map(grp => {
              const grpItems = PAGE_ACCESS_ITEMS.filter(p=>p.group===grp);
              return (
                <div key={grp} style={{marginBottom:12}}>
                  <div style={{fontSize:11,fontWeight:800,color:"var(--gold-dk)",textTransform:"uppercase",letterSpacing:.8,marginBottom:6,paddingLeft:4}}>{grp}</div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))",gap:6}}>
                    {grpItems.map(pg => {
                      const allowed = (form.allowedPages||[]).length===0 || (form.allowedPages||[]).includes(pg.key);
                      const isRestricted = (form.allowedPages||[]).length>0;
                      const checked = isRestricted ? (form.allowedPages||[]).includes(pg.key) : true;
                      return (
                        <label key={pg.key} style={{display:"flex",alignItems:"center",gap:8,cursor:isAdminRole?"default":"pointer",
                          padding:"7px 10px",borderRadius:8,
                          background:checked?"var(--gold-f)":"var(--red-f)",
                          border:`1.5px solid ${checked?"var(--gold-200)":"var(--red-200,#fecaca)"}`,
                          opacity:isAdminRole?.6:1,transition:"all .15s"}}>
                          <input type="checkbox" checked={checked} disabled={isAdminRole}
                            onChange={()=>{
                              if (isAdminRole) return;
                              const curr = form.allowedPages||[];
                              // If currently "all" (empty), switching means restrict others
                              if (curr.length===0) {
                                // Enable restriction mode: allow all EXCEPT unchecked
                                const all = PAGE_ACCESS_ITEMS.map(x=>x.key).filter(k=>k!==pg.key);
                                setForm(p=>({...p,allowedPages:all}));
                              } else if (curr.includes(pg.key)) {
                                // Remove this page
                                const next = curr.filter(k=>k!==pg.key);
                                setForm(p=>({...p,allowedPages:next}));
                              } else {
                                // Add this page
                                const next = [...curr, pg.key];
                                // If now all pages checked, reset to empty (= all)
                                if (next.length===PAGE_ACCESS_ITEMS.length) setForm(p=>({...p,allowedPages:[]}));
                                else setForm(p=>({...p,allowedPages:next}));
                              }
                            }}
                            style={{accentColor:"var(--gold)",width:15,height:15}}/>
                          <span style={{fontSize:12}}>{pg.icon}</span>
                          <span style={{fontSize:12,fontWeight:600}}>{pg.label}</span>
                        </label>
                      );
                    })}
                  </div>
                </div>
              );
            })}
            <div style={{display:"flex",gap:10,marginTop:6}}>
              <button type="button" onClick={()=>setForm(p=>({...p,allowedPages:[]}))} style={{fontSize:11,color:"var(--grn)",background:"none",border:"none",cursor:"pointer",fontWeight:700,textDecoration:"underline"}}>✅ Allow All</button>
              <button type="button" onClick={()=>setForm(p=>({...p,allowedPages:[PAGE_ACCESS_ITEMS[0].key]}))} style={{fontSize:11,color:"var(--red)",background:"none",border:"none",cursor:"pointer",fontWeight:700,textDecoration:"underline"}}>🔒 Restrict All (Dashboard only)</button>
            </div>
          </div>

          {/* Price Approval Form Field Rights */}
          <div className="fsec">
            <div className="fsec-title">💰 Price Approval Form Field Rights</div>
            <div style={{fontSize:12,color:"var(--muted)",marginBottom:10,fontWeight:500}}>
              Control which financial fields this user can <strong>view</strong> in the price approval form. Leave all unchecked = full visibility. Check specific fields = only those will be visible.
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:8,padding:6}}>
              {PRICE_FORM_FIELDS.map(f => {
                const fRights = form.formFieldRights || [];
                const restricted = fRights.length > 0;
                const hasRight = restricted ? fRights.includes(f.key) : true;
                return (
                  <label key={f.key} style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer",padding:"9px 12px",borderRadius:9,background:hasRight?"var(--gold-f)":"var(--cream)",border:"1.5px solid",borderColor:hasRight?"var(--gold-200)":"var(--border2)",transition:"all .15s"}}>
                    <input type="checkbox" style={{accentColor:"var(--gold)",width:15,height:15}} checked={hasRight}
                      onChange={()=>{
                        const curr = form.formFieldRights || [];
                        const allKeys = PRICE_FORM_FIELDS.map(x=>x.key);
                        if (curr.length === 0) {
                          // switching from "all visible" to "restricted" — exclude this field
                          setForm(p=>({...p,formFieldRights:allKeys.filter(k=>k!==f.key)}));
                        } else if (curr.includes(f.key)) {
                          const next = curr.filter(k=>k!==f.key);
                          setForm(p=>({...p,formFieldRights:next.length===0?[]:next}));
                        } else {
                          const next = [...curr, f.key];
                          setForm(p=>({...p,formFieldRights:next.length===allKeys.length?[]:next}));
                        }
                      }}
                    />
                    <span style={{fontWeight:600,fontSize:13,color:"var(--ink)"}}>{f.label}</span>
                  </label>
                );
              })}
            </div>
            <div style={{display:"flex",gap:10,marginTop:8}}>
              <button type="button" onClick={()=>setForm(p=>({...p,formFieldRights:[]}))} style={{fontSize:11,color:"var(--grn)",background:"none",border:"none",cursor:"pointer",fontWeight:700,textDecoration:"underline"}}>✅ Show All Fields</button>
              <button type="button" onClick={()=>setForm(p=>({...p,formFieldRights:["tradePrice"]}))} style={{fontSize:11,color:"var(--ora)",background:"none",border:"none",cursor:"pointer",fontWeight:700,textDecoration:"underline"}}>📋 Restrict (Trade Price only visible)</button>
            </div>
          </div>

          {/* Analytics Dashboard Rights */}
          <div className="fsec"><div className="fsec-title">📊 Analytics Dashboard Access</div>
            <div style={{fontSize:12,color:"var(--muted)",marginBottom:10,fontWeight:500}}>
              {form.role==="Admin"?"Admin has access to ALL dashboards by default.":"Select which analytics dashboards this user can access. Empty = no dashboard access."}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:8,padding:6}}>
              {ANALYTICS_DASHBOARDS.map(d => {
                const aRights = form.analyticsRights || [];
                const isAdmin = form.role === "Admin";
                const hasAccess = isAdmin || aRights.includes(d.key);
                return (
                  <label key={d.key} style={{display:"flex",alignItems:"center",gap:10,cursor:isAdmin?"default":"pointer",padding:"9px 12px",borderRadius:9,
                    background:hasAccess?"var(--gold-f)":"var(--cream)",border:"1.5px solid",
                    borderColor:hasAccess?"var(--gold-200)":"var(--border2)",transition:"all .15s",opacity:isAdmin?.7:1}}>
                    <input type="checkbox" style={{accentColor:"var(--gold)",width:15,height:15}}
                      checked={hasAccess} disabled={isAdmin}
                      onChange={()=>{
                        const curr = form.analyticsRights || [];
                        const next = curr.includes(d.key) ? curr.filter(k=>k!==d.key) : [...curr, d.key];
                        setForm(p=>({...p, analyticsRights: next}));
                      }}
                    />
                    <div>
                      <div style={{fontWeight:700,fontSize:12}}>{d.icon} {d.label}</div>
                      <div style={{fontSize:10.5,color:"var(--muted)"}}>{d.desc}</div>
                    </div>
                  </label>
                );
              })}
            </div>
            <div style={{display:"flex",gap:10,marginTop:8}}>
              <button type="button" onClick={()=>setForm(p=>({...p,analyticsRights:ANALYTICS_DASHBOARDS.map(d=>d.key)}))} style={{fontSize:11,color:"var(--grn)",background:"none",border:"none",cursor:"pointer",fontWeight:700,textDecoration:"underline"}}>✅ Grant All Dashboards</button>
              <button type="button" onClick={()=>setForm(p=>({...p,analyticsRights:[]}))} style={{fontSize:11,color:"var(--red)",background:"none",border:"none",cursor:"pointer",fontWeight:700,textDecoration:"underline"}}>🚫 Revoke All</button>
            </div>
          </div>

          {/* Customer Assignment */}
          <div className="fsec"><div className="fsec-title">🔗 Customer Assignment</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:8,maxHeight:180,overflowY:"auto",padding:6,border:"1.5px solid var(--border2)",borderRadius:10,background:"var(--cream)"}}>
              {customers.map(c=><label key={c.customerCode} style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer",padding:"9px 12px",borderRadius:9,background:form.assignedCustomers.includes(c.customerCode)?"var(--gold-f)":"var(--white)",border:"1.5px solid",borderColor:form.assignedCustomers.includes(c.customerCode)?"var(--gold-200)":"var(--border2)",transition:"all .15s"}}><input type="checkbox" checked={form.assignedCustomers.includes(c.customerCode)} onChange={()=>toggleCust(c.customerCode)} style={{accentColor:"var(--gold)",width:16,height:16}}/><div><div style={{fontWeight:700,fontSize:13}}>{c.customerName}</div><div style={{fontSize:11,color:"var(--muted)"}}>{c.customerCode}</div></div></label>)}
            </div>
          </div>
        </div>
        <div className="m-footer"><button className="btn btn-secondary btn-sm" onClick={onClose}>Cancel</button><button className="btn btn-primary btn-sm" onClick={()=>{if(validate())onSave(form);}}>{isEdit?"💾 Update":"✅ Create"}</button></div>
      </div>
    </div>
  );
}

function UserMasterPanel({ users, setUsers, customers, locationMaster }) {
  const [modal, setModal] = useState(null); const [search, setSearch] = useState(""); const [fRole, setFRole] = useState(""); const [fStatus, setFStatus] = useState("");
  const { exportCsv, exportTemplate, ExportModal, ToastEl } = useExport();
  const [toast2, showToast2] = useToast();
  const filtered = users.filter(u => { if (fRole && u.role !== fRole) return false; if (fStatus && (u.status || "Active") !== fStatus) return false; if (search) { const q = search.toLowerCase(); return u.name.toLowerCase().includes(q) || u.id.toLowerCase().includes(q) || u.email.toLowerCase().includes(q); } return true; });
  const handleSave = data => { setUsers(prev => { const idx = prev.findIndex(u => u.id === data.id); if (idx >= 0) { const n = [...prev]; n[idx] = data; return n; } return [...prev, data]; }); showToast2(data.name + " saved!"); setModal(null); };
  const handleDelete = id => { setUsers(prev => prev.filter(u => u.id !== id)); showToast2("Deleted.", "danger"); };
  const handlePwChange = (uid, pw) => { setUsers(prev => prev.map(u => u.id === uid ? { ...u, password: pw } : u)); showToast2("Password updated!"); };
  const handleImport = recs => { setUsers(prev => { const map = new Map(prev.map(u => [u.id, u])); recs.forEach(r => { if (r.id) map.set(r.id, { ...map.get(r.id) || { password: "changeme", assignedCustomers: [] }, ...r }); }); return [...map.values()]; }); showToast2("Imported!"); };
  return (
    <div>
      {toast2 && <Toast msg={toast2.msg} type={toast2.type} />}
      {ToastEl}
      <div className="card">
        <div className="card-hdr">
          <div><div className="card-title">👤 User Management</div><div className="card-sub">{users.length} users</div></div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
            <div className="si-wrap" style={{ minWidth: 180 }}><span className="si-ico">🔍</span><input className="si" placeholder="Search…" value={search} onChange={e => setSearch(e.target.value)} /></div>
            <select className="fsel" value={fRole} onChange={e => setFRole(e.target.value)}><option value="">All Roles</option>{ROLES.map(r => <option key={r}>{r}</option>)}</select>
            <select className="fsel" value={fStatus} onChange={e => setFStatus(e.target.value)}><option value="">All Status</option><option>Active</option><option>Inactive</option></select>
            {(search || fRole || fStatus) && <button className="btn btn-ghost btn-sm" onClick={() => { setSearch(""); setFRole(""); setFStatus(""); }}>✕</button>}
            <button className="btn btn-secondary btn-sm" onClick={() => exportTemplate(USER_COLS, "user_template.csv")}>📋 Template</button>
            <button className="btn btn-secondary btn-sm" onClick={() => setModal({ type: "import" })}>📤 Import</button>
            <button className="btn btn-secondary btn-sm" onClick={() => exportCsv(filtered.map(u => ({ ...u, password: "***" })), USER_COLS, `users_${new Date().toISOString().split("T")[0]}.csv`, "Users (passwords masked)")}>⬇ Export</button>
            <button className="btn btn-primary btn-sm" onClick={() => setModal({ type: "form", data: null })}>+ Add</button>
          </div>
        </div>
        <div className="tw"><table>
          <thead><tr><th>ID</th><th>Name</th><th>Role</th><th>Email</th><th>Status</th><th>Customers</th><th>Actions</th></tr></thead>
          <tbody>
            {filtered.length === 0 && <tr><td colSpan={7} style={{ textAlign: "center", padding: "40px", color: "var(--muted)" }}>No users found</td></tr>}
            {filtered.map(u => (
              <tr key={u.id}>
                <td><span style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: "var(--gold-dk)", fontSize: 12.5 }}>{u.id}</span></td>
                <td><div style={{ fontWeight: 700, fontSize: 13 }}>{u.name}</div>{u.designation && <div style={{ fontSize: 11, color: "var(--muted)" }}>{u.designation}</div>}</td>
                <td><span className={`badge ${ROLE_BADGE[u.role] || "b-so"}`}>{u.role}</span></td>
                <td style={{ fontSize: 12.5, color: "var(--ink3)" }}>{u.email}</td>
                <td><span className={`badge ${(u.status || "Active") === "Active" ? "b-active" : "b-inactive"}`}>{u.status || "Active"}</span></td>
                <td>
  {u.role==="Admin"||u.role==="Zonal Head"||u.role==="Sales & Accounts"
    ? <span style={{fontSize:12,color:"var(--muted)",fontStyle:"italic"}}>All (by role)</span>
    : (() => {
        const locs = ["allowedSalesOffices","allowedRegions","allowedDistricts","allowedTehsils"].flatMap(k=>(u[k]||[]).map(id=>id));
        const custs = (u.assignedCustomers||[]);
        return (
          <div style={{display:"flex",flexDirection:"column",gap:3}}>
            {custs.length>0&&<div style={{display:"flex",gap:3,flexWrap:"wrap"}}>{custs.slice(0,2).map(c=><span key={c} className="badge b-grade">{c}</span>)}{custs.length>2&&<span style={{fontSize:11,color:"var(--muted)"}}>+{custs.length-2}</span>}</div>}
            {locs.length>0&&<span style={{fontSize:11,color:"var(--gold-dk)",fontWeight:700}}>📍 {locs.length} loc rights</span>}
            {(u.allowedPages||[]).length>0&&<span style={{fontSize:11,color:"var(--red)",fontWeight:700}}>🔐 {u.allowedPages.length} pages</span>}
            {locs.length===0&&custs.length===0&&<span style={{fontSize:11,color:"var(--muted)",fontStyle:"italic"}}>All</span>}
          </div>
        );
      })()
  }
</td>
                <td><div style={{ display: "flex", gap: 3 }}><button className="btn btn-xs btn-ghost" onClick={() => setModal({ type: "form", data: u })}>✏️</button><button className="btn btn-xs btn-ghost" onClick={() => setModal({ type: "pw", data: u })}>🔑</button><button className="btn btn-xs btn-ghost" style={{ color: "var(--red)" }} onClick={() => setModal({ type: "delete", data: u })}>🗑</button></div></td>
              </tr>
            ))}
          </tbody>
        </table></div>
        <div style={{ padding: "10px 20px", borderTop: "1.5px solid var(--border2)", fontSize: 12, color: "var(--muted)" }}>💡 Click 🔑 to change passwords. Passwords masked in exports.</div>
      </div>
      {modal?.type === "form" && <UserFormModal user={modal.data} allUsers={users} customers={customers} locationMaster={locationMaster} onSave={handleSave} onClose={() => setModal(null)} />}
      {modal?.type === "pw" && <PwModal user={modal.data} onSave={pw => handlePwChange(modal.data.id, pw)} onClose={() => setModal(null)} />}
      {modal?.type === "delete" && <ConfirmDelete label={`${modal.data.name} (${modal.data.id})`} onConfirm={() => { handleDelete(modal.data.id); setModal(null); }} onClose={() => setModal(null)} />}
      {modal?.type === "import" && <ImportModal title="Users" colDefs={USER_COLS} previewCols={["id", "name", "role", "email", "status"]} onImport={handleImport} onClose={() => setModal(null)} />}
      {ExportModal}
    </div>
  );
}


// ══════════════════════════════════════════════════════
// SALES ENTRY MODAL
// ══════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════
// SALE UPDATION PAGE
// Cols match Sale_Updation.xlsx exactly
// ══════════════════════════════════════════════════════
function SaleUpdationPage({ requests, salesEntries, setSalesEntries, user }) {
  const emptyForm = { requestId:"", customerCode:"", customerName:"", billingDocNo:"", billingDocDate:"", materialGrade:"OPC", invoiceQty:"", invoiceValue:"", ratePerBag:"" };
  const [form, setForm] = useState(emptyForm);
  const [editId, setEditId] = useState(null);
  const [errors, setErrors] = useState({});
  const [search, setSearch] = useState("");
  const fileRef = useRef();

  // When requestId changes, auto-fill customer info
  const updateReq = (reqId) => {
    const req = requests.find(r=>r.id===reqId);
    setForm(p=>({ ...p, requestId:reqId,
      customerCode: req?.customerCode||"",
      customerName: req?.customerName||"",
      materialGrade: req?.grade||p.materialGrade,
    }));
  };

  const validate = () => {
    const e = {};
    if (!form.requestId)     e.requestId   = "Required";
    if (!form.billingDocNo.trim()) e.billingDocNo = "Required";
    if (!form.billingDocDate) e.billingDocDate = "Required";
    if (!form.invoiceQty || isNaN(Number(form.invoiceQty)))   e.invoiceQty  = "Required";
    if (!form.invoiceValue|| isNaN(Number(form.invoiceValue))) e.invoiceValue= "Required";
    if (!form.ratePerBag  || isNaN(Number(form.ratePerBag)))  e.ratePerBag  = "Required";
    setErrors(e); return !Object.keys(e).length;
  };

  const handleSave = () => {
    if (!validate()) return;
    setSalesEntries(prev => {
      if (editId) return prev.map(e=>e.id===editId ? {...e,...form,enteredBy:user?.id,enteredAt:new Date().toISOString()} : e);
      return [...prev, {id:`SE${Date.now()}`,...form,enteredBy:user?.id,enteredAt:new Date().toISOString()}];
    });
    setForm(emptyForm); setEditId(null); setErrors({});
  };

  const handleEdit = (e) => { setForm({requestId:e.requestId,customerCode:e.customerCode,customerName:e.customerName,billingDocNo:e.billingDocNo,billingDocDate:e.billingDocDate,materialGrade:e.materialGrade,invoiceQty:e.invoiceQty,invoiceValue:e.invoiceValue,ratePerBag:e.ratePerBag}); setEditId(e.id); };
  const handleDelete = (id) => { if(window.confirm("Delete this entry?")) setSalesEntries(p=>p.filter(e=>e.id!==id)); };

  // Excel upload matching Sale_Updation.xlsx
  const handleFile = (ev) => {
    const file = ev.target.files[0]; if(!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result,{type:"array"});
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws,{defval:""});
      const newEntries = rows.filter(r=>r["Price Report ID"]).map(r=>({
        id:`SE${Date.now()}_${Math.random().toString(36).slice(2)}`,
        requestId:     String(r["Price Report ID"]||""),
        customerCode:  String(r["Customer code"]||""),
        customerName:  String(r["Customer name"]||""),
        billingDocNo:  String(r["Billing Document No."]||""),
        billingDocDate:String(r["Billing Document Date"]||""),
        materialGrade: String(r["Material Grade"]||""),
        invoiceQty:    String(r["Invoice Qty."]||""),
        invoiceValue:  String(r["Invoice Value"]||""),
        ratePerBag:    String(r["Rate per Bag"]||""),
        enteredBy:     user?.id, enteredAt:new Date().toISOString(),
      }));
      setSalesEntries(p=>[...p,...newEntries]);
      alert(`Imported ${newEntries.length} entries.`);
    };
    reader.readAsArrayBuffer(file); ev.target.value="";
  };

  // Excel export matching Sale_Updation.xlsx
  const handleExport = () => {
    const hdr = ["Price Report ID","Customer code","Customer name","Billing Document No.","Billing Document Date","Material Grade","Invoice Qty.","Invoice Value","Rate per Bag"];
    const rows = filtered.map(e=>[e.requestId,e.customerCode,e.customerName,e.billingDocNo,e.billingDocDate,e.materialGrade,e.invoiceQty,e.invoiceValue,e.ratePerBag]);
    const ws = XLSX.utils.aoa_to_sheet([hdr,...rows]);
    ws["!cols"] = hdr.map(()=>({wch:18}));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Sale Updation");
    XLSX.writeFile(wb,`Sale_Updation_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  const handleTemplate = () => {
    const hdr = ["Price Report ID","Customer code","Customer name","Billing Document No.","Billing Document Date","Material Grade","Invoice Qty.","Invoice Value","Rate per Bag"];
    const sample = [["REQ-001","C001","Raj Construction Co.","4900001234","2024-04-01","PPC",50,150000,330]];
    const ws = XLSX.utils.aoa_to_sheet([hdr,...sample]);
    ws["!cols"] = hdr.map(()=>({wch:20}));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Sale Updation");
    XLSX.writeFile(wb,"Sale_Updation_Template.xlsx");
  };

  const filtered = salesEntries.filter(e =>
    !search || e.requestId.toLowerCase().includes(search.toLowerCase()) ||
    e.customerName.toLowerCase().includes(search.toLowerCase()) ||
    e.billingDocNo.toLowerCase().includes(search.toLowerCase())
  );

  const Err = ({k}) => errors[k]?<span style={{fontSize:11,color:"var(--red)",fontWeight:600}}>{errors[k]}</span>:null;
  const THD = ({children,right=false}) => <th style={{padding:"9px 12px",background:"var(--navy-mid)",color:"var(--gold-lt)",fontWeight:700,fontSize:11,textTransform:"uppercase",letterSpacing:.5,textAlign:right?"right":"left",whiteSpace:"nowrap"}}>{children}</th>;
  const TD  = ({children,right=false,bold=false,col=undefined}) => <td style={{padding:"8px 12px",borderBottom:"1px solid var(--border2)",textAlign:right?"right":"left",fontWeight:bold?700:400,color:col}}>{children}</td>;

  return (
    <div>
      <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={handleFile}/>
      <div className="ph">
        <div><div className="ph-eyebrow">Sales</div><div className="ph-title">Sale Updation</div>
          <div className="ph-sub">Record actual billing data against approved price requests</div>
        </div>
        <div className="ph-actions">
          <button className="btn btn-secondary btn-sm" onClick={handleTemplate}>📋 Template</button>
          <button className="btn btn-secondary btn-sm" onClick={()=>fileRef.current?.click()}>📤 Upload Excel</button>
          <button className="btn btn-primary btn-sm" onClick={handleExport}>⬇ Export Excel</button>
        </div>
      </div>

      {/* KPIs */}
      <div className="kpi-grid" style={{marginBottom:16}}>
        {[
          {l:"Total Entries",    v:salesEntries.length, i:"📋",c:"var(--gold)"},
          {l:"Price Requests Covered", v:new Set(salesEntries.map(e=>e.requestId)).size, i:"🔗",c:"#7C3AED"},
          {l:"Total Invoice Qty MT", v:salesEntries.reduce((s,e)=>s+Number(e.invoiceQty||0),0).toLocaleString("en-IN"), i:"📦",c:"var(--grn)"},
          {l:"Total Invoice Value", v:`₹${(salesEntries.reduce((s,e)=>s+Number(e.invoiceValue||0),0)/100000).toFixed(1)}L`, i:"💰",c:"var(--grn)"},
        ].map((k,i)=>(
          <div key={i} className="kpi"><div className="kpi-accent" style={{background:k.c}}/><div className="kpi-icon">{k.i}</div>
            <div className="kpi-val" style={{color:k.c,fontSize:k.v.toString().length>8?16:24}}>{k.v}</div><div className="kpi-label">{k.l}</div>
          </div>
        ))}
      </div>

      {/* Entry Form */}
      <div className="card" style={{marginBottom:16}}>
        <div className="card-hdr"><div className="card-title">{editId?"✏️ Edit Entry":"➕ New Sale Entry"}</div></div>
        <div className="card-body">
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:12,marginBottom:12}}>
            <div className="fg">
              <label className="fl fl-req">Price Report ID</label>
              <select className="fc" value={form.requestId} onChange={e=>{updateReq(e.target.value);}}>
                <option value="">-- Select Request --</option>
                {[...new Set([...requests.map(r=>r.id)])].map(id=>{const req=requests.find(r=>r.id===id); return <option key={id} value={id}>{id} — {req?.customerName||""} ({req?.grade||""})</option>;})}
              </select><Err k="requestId"/>
            </div>
            <div className="fg"><label className="fl">Customer Code</label><input className="fc" value={form.customerCode} readOnly style={{background:"var(--cream)",color:"var(--muted)"}}/></div>
            <div className="fg"><label className="fl">Customer Name</label><input className="fc" value={form.customerName} readOnly style={{background:"var(--cream)",color:"var(--muted)"}}/></div>
            <div className="fg"><label className="fl fl-req">Billing Doc No.</label><input className="fc" value={form.billingDocNo} onChange={e=>setForm(p=>({...p,billingDocNo:e.target.value}))}/><Err k="billingDocNo"/></div>
            <div className="fg"><label className="fl fl-req">Billing Doc Date</label><input className="fc" type="date" value={form.billingDocDate} onChange={e=>setForm(p=>({...p,billingDocDate:e.target.value}))}/><Err k="billingDocDate"/></div>
            <div className="fg"><label className="fl">Material Grade</label><select className="fc" value={form.materialGrade} onChange={e=>setForm(p=>({...p,materialGrade:e.target.value}))}><option>OPC</option><option>PPC</option><option>ProMaxX</option></select></div>
            <div className="fg"><label className="fl fl-req">Invoice Qty (MT)</label><input className="fc" type="number" value={form.invoiceQty} onChange={e=>setForm(p=>({...p,invoiceQty:e.target.value}))}/><Err k="invoiceQty"/></div>
            <div className="fg"><label className="fl fl-req">Invoice Value (₹)</label><input className="fc" type="number" value={form.invoiceValue} onChange={e=>setForm(p=>({...p,invoiceValue:e.target.value}))}/><Err k="invoiceValue"/></div>
            <div className="fg"><label className="fl fl-req">Rate per Bag (₹)</label><input className="fc" type="number" value={form.ratePerBag} onChange={e=>setForm(p=>({...p,ratePerBag:e.target.value}))}/><Err k="ratePerBag"/></div>
          </div>
          <div style={{display:"flex",gap:10}}>
            <button className="btn btn-primary btn-sm" onClick={handleSave}>{editId?"💾 Update":"✅ Save Entry"}</button>
            {editId && <button className="btn btn-ghost btn-sm" onClick={()=>{setForm(emptyForm);setEditId(null);}}>✕ Cancel</button>}
          </div>
        </div>
      </div>

      {/* Records Table */}
      <div className="card">
        <div className="card-hdr">
          <div className="card-title">📋 Sale Entries ({filtered.length})</div>
          <input className="fc" placeholder="🔍 Search by Request ID / Customer / Doc No..." value={search} onChange={e=>setSearch(e.target.value)} style={{maxWidth:320,padding:"6px 12px",fontSize:13}}/>
        </div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr>
              <THD>Price Report ID</THD><THD>Customer Code</THD><THD>Customer Name</THD>
              <THD>Billing Doc No.</THD><THD>Billing Date</THD><THD>Grade</THD>
              <THD right>Invoice Qty (MT)</THD><THD right>Invoice Value (₹)</THD><THD right>Rate/Bag (₹)</THD>
              <THD>Entered By</THD><THD>Actions</THD>
            </tr></thead>
            <tbody>
              {filtered.length===0 && <tr><td colSpan={11} style={{textAlign:"center",padding:40,color:"var(--muted)"}}>No entries yet. Add manually or upload Excel.</td></tr>}
              {filtered.map((e,i)=>(
                <tr key={e.id} style={{background:i%2===0?"var(--white)":"var(--cream)"}}>
                  <TD bold col="var(--gold-dk)">{e.requestId}</TD>
                  <TD col="var(--muted)">{e.customerCode}</TD>
                  <TD bold>{e.customerName}</TD>
                  <TD>{e.billingDocNo}</TD>
                  <TD col="var(--muted)">{e.billingDocDate}</TD>
                  <TD><span className="badge b-grade">{e.materialGrade}</span></TD>
                  <TD right bold>{Number(e.invoiceQty||0).toLocaleString("en-IN")}</TD>
                  <TD right bold col="var(--grn)">₹{Number(e.invoiceValue||0).toLocaleString("en-IN")}</TD>
                  <TD right bold>₹{Number(e.ratePerBag||0).toLocaleString("en-IN")}</TD>
                  <TD col="var(--muted)">{e.enteredBy}</TD>
                  <TD><div style={{display:"flex",gap:6}}>
                    <button onClick={()=>handleEdit(e)} style={{fontSize:16,background:"none",border:"none",cursor:"pointer"}}>✏️</button>
                    <button onClick={()=>handleDelete(e.id)} style={{fontSize:16,background:"none",border:"none",cursor:"pointer"}}>🗑</button>
                  </div></TD>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════
// SALES REPORT PAGE
// Cols match Sales_Report.xlsb exactly
// Price Request + Sales Entries combined analysis
// ══════════════════════════════════════════════════════
function SalesReportPage({ requests, salesEntries, user }) {
  const [fGrade,setFGrade]     = useState("");
  const [fRegion,setFRegion]   = useState("");
  const [fStatus,setFStatus]   = useState("Approved");
  const [dateFrom,setDateFrom] = useState(""); const [dateTo,setDateTo] = useState("");
  const [view,setView]         = useState("table");

  const fReqs = requests.filter(r => {
    if (fGrade  && r.grade  !== fGrade)  return false;
    if (fRegion && r.region !== fRegion) return false;
    if (fStatus && r.status !== fStatus) return false;
    if (dateFrom && r.date < dateFrom)   return false;
    if (dateTo   && r.date > dateTo)     return false;
    return true;
  });

  const enriched = fReqs.map(req => {
    const entries      = (salesEntries||[]).filter(e => e.requestId === req.id);
    const totalInvQty  = entries.reduce((s,e)=>s+Number(e.invoiceQty||0),0);
    const totalInvVal  = entries.reduce((s,e)=>s+Number(e.invoiceValue||0),0);
    const avgRatePerBag= entries.length>0 ? entries.reduce((s,e)=>s+Number(e.ratePerBag||0),0)/entries.length : 0;
    const achievePct   = Number(req.qty||0)>0 ? Math.round(totalInvQty/Number(req.qty)*100) : 0;
    const rateVariance = avgRatePerBag && Number(req.orderPrice||0) ? avgRatePerBag-Number(req.orderPrice) : null;
    return {...req, entries, totalInvQty, totalInvVal, avgRatePerBag, achievePct, rateVariance};
  });

  const totalReqQty = enriched.reduce((s,r)=>s+Number(r.qty||0),0);
  const totalInvQty = enriched.reduce((s,r)=>s+r.totalInvQty,0);
  const totalInvVal = enriched.reduce((s,r)=>s+r.totalInvVal,0);
  const overallAch  = totalReqQty>0 ? Math.round(totalInvQty/totalReqQty*100) : 0;
  const regions     = [...new Set(requests.map(r=>r.region).filter(Boolean))];
  const ach = (p) => p>=100?"var(--grn)":p>=75?"var(--gold)":"var(--red)";
  const fmt = (v) => v!=null&&v!==""?Number(v).toLocaleString("en-IN"):"—";

  // Export — all form fields + sales entry cols (matches Sales_Report.xlsb + full expenses)
  const handleExport = () => {
    const hdr = [
      // Price Request fields
      "Request ID","Date","Customer Code","Customer Name","Region","Zone","Destination",
      "Grade","Material Type","Qty (MT)","Order Price (₹)","Trade Price (₹)","Difference (₹)",
      "Net of GST (₹)","Payment Terms","Supply From","Valid From","Valid To","TPC Agent",
      // Expenses
      "Primary Freight","Secondary Freight","Demurrage","Station Handling",
      "Cost of Production","Packing","Unloading","OP Commission","SP Commission",
      "Total Expenses","NCR/MT",
      // Approval
      "Status",
      // Actual Sales
      "Billing Document No.","Billing Document Date","Material Grade",
      "Invoice Qty. (MT)","Invoice Value (₹)","Rate per Bag (₹)",
      // Performance
      "Achievement %","Rate Variance (₹)"
    ];
    const rows = enriched.flatMap(r => {
      const base = [
        r.id, r.date, r.customerCode, r.customerName, r.region||"", r.zone||"", r.destination||"",
        r.grade, r.materialType, r.qty, r.orderPrice, r.tradePrice, r.difference,
        r.netOfGST, r.paymentTerms, r.supplyFrom, r.validityFrom, r.validityTo, r.tpcAgent||"",
        r.primaryFreight||"", r.secondaryFreight||"", r.demrage||"", r.stationHandling||"",
        r.costOfProduction||"", r.packing||"", r.unloadingPrice||"", r.opCommission||"", r.spCommission||"",
        r.totalExpenses||"", r.ncrPmt||"",
        r.status
      ];
      if (r.entries.length===0) return [[...base,"","","","","","","",""]];
      return r.entries.map(e=>[...base,
        e.billingDocNo, e.billingDocDate, e.materialGrade, e.invoiceQty, e.invoiceValue, e.ratePerBag,
        r.achievePct+"%", r.rateVariance!=null?r.rateVariance.toFixed(2):""
      ]);
    });
    const ws = XLSX.utils.aoa_to_sheet([hdr,...rows]);
    ws["!cols"] = hdr.map(()=>({wch:18}));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Sales Report");
    XLSX.writeFile(wb,`Sales_Report_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  // Column group helper
  const GHD = ({children,span,bg}) => <th colSpan={span} style={{padding:"6px 10px",background:bg||"var(--navy)",color:"var(--gold-lt)",fontWeight:800,fontSize:10,textTransform:"uppercase",letterSpacing:.6,textAlign:"center",borderRight:"2px solid rgba(255,255,255,.1)"}}>{children}</th>;
  const THD = ({children,right=false}) => <th style={{padding:"7px 8px",background:"var(--navy-mid)",color:"var(--gold-lt)",fontWeight:700,fontSize:10,textTransform:"uppercase",letterSpacing:.3,textAlign:right?"right":"left",whiteSpace:"nowrap"}}>{children}</th>;
  const TD  = ({children,right=false,bold=false,col=undefined,small=false}) => <td style={{padding:"6px 8px",borderBottom:"1px solid var(--border2)",textAlign:right?"right":"left",fontWeight:bold?700:400,color:col,fontSize:small?11:12,whiteSpace:"nowrap"}}>{children}</td>;

  return (
    <div>
      <div className="ph">
        <div><div className="ph-eyebrow">Analytics</div><div className="ph-title">Sales Report</div>
          <div className="ph-sub">All price request components + actual sales — full performance analysis</div>
        </div>
        <div className="ph-actions">
          <button className="btn btn-secondary btn-sm" onClick={()=>setView(v=>v==="table"?"summary":"table")}>{view==="table"?"📊 Summary":"📋 Detailed"}</button>
          <button className="btn btn-primary btn-sm" onClick={handleExport}>⬇ Export Excel</button>
        </div>
      </div>

      {/* KPIs */}
      <div className="kpi-grid">
        {[
          {l:"Price Requests",   v:enriched.length,                        i:"📋",c:"var(--gold)"},
          {l:"With Sales Data",  v:enriched.filter(r=>r.entries.length>0).length,i:"✅",c:"var(--grn)"},
          {l:"Req. Qty MT",      v:totalReqQty.toLocaleString("en-IN"),    i:"📦",c:"#7C3AED"},
          {l:"Invoice Qty MT",   v:totalInvQty.toLocaleString("en-IN"),    i:"🧾",c:"var(--grn)"},
          {l:"Achievement",      v:`${overallAch}%`,                       i:"🎯",c:ach(overallAch)},
          {l:"Invoice Value",    v:`₹${(totalInvVal/100000).toFixed(1)}L`, i:"💰",c:"var(--grn)"},
        ].map((k,i)=>(
          <div key={i} className="kpi"><div className="kpi-accent" style={{background:k.c}}/><div className="kpi-icon">{k.i}</div>
            <div className="kpi-val" style={{color:k.c,fontSize:k.v.toString().length>8?16:24}}>{k.v}</div>
            <div className="kpi-label">{k.l}</div>
          </div>
        ))}
      </div>

      {/* Filters */}
      <div className="card" style={{marginBottom:16}}>
        <div className="card-hdr"><div className="card-title">🔍 Filter</div></div>
        <div className="card-body"><div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(145px,1fr))",gap:12}}>
          <div className="fg"><label className="fl">Status</label><select className="fc" value={fStatus} onChange={e=>setFStatus(e.target.value)}><option value="">All</option><option>Approved</option><option>Pending</option><option>Rejected</option></select></div>
          <div className="fg"><label className="fl">Grade</label><select className="fc" value={fGrade} onChange={e=>setFGrade(e.target.value)}><option value="">All</option><option>OPC</option><option>PPC</option><option>ProMaxX</option></select></div>
          <div className="fg"><label className="fl">Region</label><select className="fc" value={fRegion} onChange={e=>setFRegion(e.target.value)}><option value="">All</option>{regions.map(r=><option key={r}>{r}</option>)}</select></div>
          <div className="fg"><label className="fl">From</label><input className="fc" type="date" value={dateFrom} onChange={e=>setDateFrom(e.target.value)}/></div>
          <div className="fg"><label className="fl">To</label><input className="fc" type="date" value={dateTo} onChange={e=>setDateTo(e.target.value)}/></div>
        </div></div>
      </div>

      {view==="table" ? (
        <div className="card">
          <div className="card-hdr"><div className="card-title">Complete Sales Report — All Components</div><span style={{fontSize:11,color:"var(--muted)"}}>← scroll horizontally to see all columns →</span></div>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",minWidth:2200}}>
              <thead>
                {/* Column group headers */}
                <tr>
                  <GHD span={7}  bg="var(--navy)">📋 Request Info</GHD>
                  <GHD span={5}  bg="#1e3a5f">💰 Pricing</GHD>
                  <GHD span={3}  bg="#1a3320">📦 Order</GHD>
                  <GHD span={9}  bg="#3d1a00">🏭 Expenses</GHD>
                  <GHD span={2}  bg="#0d2233">📊 Summary</GHD>
                  <GHD span={1}  bg="#1a0d33">Status</GHD>
                  <GHD span={6}  bg="#0d3320">🧾 Actual Sales</GHD>
                  <GHD span={2}  bg="#1a3300">🎯 Performance</GHD>
                </tr>
                <tr>
                  {/* Request Info */}
                  <THD>Req ID</THD><THD>Date</THD><THD>Cust Code</THD><THD>Customer</THD><THD>Region</THD><THD>Zone</THD><THD>Destination</THD>
                  {/* Pricing */}
                  <THD right>Order Price ₹</THD><THD right>Trade Price ₹</THD><THD right>Difference ₹</THD><THD right>Net of GST ₹</THD><THD>Payment</THD>
                  {/* Order */}
                  <THD>Grade</THD><THD>Material</THD><THD right>Qty MT</THD>
                  {/* Expenses */}
                  <THD right>Pri.Freight</THD><THD right>Sec.Freight</THD><THD right>Demurrage</THD><THD right>Stn.Handling</THD><THD right>CoP</THD><THD right>Packing</THD><THD right>Unloading</THD><THD right>OP Comm.</THD><THD right>SP Comm.</THD>
                  {/* Summary */}
                  <THD right>Total Exp.</THD><THD right>NCR/MT ₹</THD>
                  {/* Status */}
                  <THD>Status</THD>
                  {/* Actual Sales */}
                  <THD>Billing Doc</THD><THD>Bill Date</THD><THD>Grade</THD><THD right>Inv.Qty MT</THD><THD right>Inv.Value ₹</THD><THD right>Rate/Bag ₹</THD>
                  {/* Performance */}
                  <THD right>Achieve%</THD><THD right>Rate Var.</THD>
                </tr>
              </thead>
              <tbody>
                {enriched.length===0 && <tr><td colSpan={35} style={{textAlign:"center",padding:40,color:"var(--muted)"}}>No data. Adjust filters.</td></tr>}
                {enriched.map((r,i) => {
                  const rows = r.entries.length>0 ? r.entries : [null];
                  return rows.map((e,j) => (
                    <tr key={`${r.id}-${j}`} style={{background:i%2===0?"var(--white)":"var(--cream)"}}>
                      {/* Request Info — only on first row */}
                      {j===0 ? <>
                        <TD bold col="var(--gold-dk)">{r.id}</TD>
                        <TD small col="var(--muted)">{r.date}</TD>
                        <TD small col="var(--muted)">{r.customerCode}</TD>
                        <TD bold>{r.customerName}</TD>
                        <TD small>{r.region||"—"}</TD>
                        <TD small>{r.zone||"—"}</TD>
                        <TD small>{r.destination||"—"}</TD>
                        {/* Pricing */}
                        <TD right bold>₹{fmt(r.orderPrice)}</TD>
                        <TD right col="var(--muted)">₹{fmt(r.tradePrice)}</TD>
                        <TD right col={Number(r.difference||0)<0?"var(--red)":"var(--grn)"}>{fmt(r.difference)}</TD>
                        <TD right>₹{fmt(r.netOfGST)}</TD>
                        <TD small>{r.paymentTerms||"—"}</TD>
                        {/* Order */}
                        <TD><span className="badge b-grade">{r.grade}</span></TD>
                        <TD small>{r.materialType||"—"}</TD>
                        <TD right bold>{fmt(r.qty)}</TD>
                        {/* Expenses */}
                        <TD right small>{fmt(r.primaryFreight)}</TD>
                        <TD right small>{fmt(r.secondaryFreight)}</TD>
                        <TD right small>{fmt(r.demrage)}</TD>
                        <TD right small>{fmt(r.stationHandling)}</TD>
                        <TD right small>{fmt(r.costOfProduction)}</TD>
                        <TD right small>{fmt(r.packing)}</TD>
                        <TD right small>{fmt(r.unloadingPrice)}</TD>
                        <TD right small>{fmt(r.opCommission)}</TD>
                        <TD right small>{fmt(r.spCommission)}</TD>
                        {/* Summary */}
                        <TD right bold col="var(--gold-dk)">₹{fmt(r.totalExpenses)}</TD>
                        <TD right bold col={Number(r.ncrPmt||0)<0?"var(--red)":"var(--grn)"}>₹{fmt(r.ncrPmt)}</TD>
                        {/* Status */}
                        <TD><span className={`badge ${r.status==="Approved"?"b-approved":r.status==="Rejected"?"b-rejected":"b-pending"}`}>{r.status}</span></TD>
                      </> : <>{Array(23).fill(null).map((_,k)=><td key={k} style={{padding:"6px 8px",borderBottom:"1px solid var(--border2)"}}/>)}</>}
                      {/* Sales entry row */}
                      {e ? <>
                        <TD>{e.billingDocNo}</TD>
                        <TD small col="var(--muted)">{e.billingDocDate}</TD>
                        <TD><span className="badge b-grade">{e.materialGrade}</span></TD>
                        <TD right bold>{Number(e.invoiceQty||0).toLocaleString("en-IN")}</TD>
                        <TD right col="var(--grn)">₹{Number(e.invoiceValue||0).toLocaleString("en-IN")}</TD>
                        <TD right bold>₹{Number(e.ratePerBag||0).toLocaleString("en-IN")}</TD>
                      </> : <>
                        <TD col="var(--muted)" small>No sales data</TD>
                        {Array(5).fill(null).map((_,k)=><td key={k} style={{padding:"6px 8px",borderBottom:"1px solid var(--border2)"}}/>)}
                      </>}
                      {/* Performance — first row only */}
                      {j===0 ? <>
                        <TD right><span style={{fontWeight:900,color:r.entries.length>0?ach(r.achievePct):"var(--muted)"}}>{r.entries.length>0?`${r.achievePct}%`:"—"}</span></TD>
                        <TD right><span style={{fontWeight:700,color:r.rateVariance!=null?(r.rateVariance>=0?"var(--grn)":"var(--red)"):"var(--muted)"}}>{r.rateVariance!=null?(r.rateVariance>=0?"+":"")+r.rateVariance.toFixed(2):"—"}</span></TD>
                      </> : <>{Array(2).fill(null).map((_,k)=><td key={k} style={{padding:"6px 8px",borderBottom:"1px solid var(--border2)"}}/>)}</>}
                    </tr>
                  ));
                })}
              </tbody>
            </table>
          </div>
        </div>
      ) : (
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
          <div className="card"><div className="card-hdr"><div className="card-title">Grade-wise Achievement</div></div>
            <div className="card-body">
              {["OPC","PPC","ProMaxX"].map(g=>{
                const gr=enriched.filter(r=>r.grade===g);
                const rQ=gr.reduce((s,r)=>s+Number(r.qty||0),0);
                const iQ=gr.reduce((s,r)=>s+r.totalInvQty,0);
                const pct=rQ>0?Math.round(iQ/rQ*100):0;
                return <div key={g} style={{marginBottom:14}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}><b>{g}</b><span style={{color:ach(pct),fontWeight:800}}>{pct}% · {iQ.toLocaleString()} / {rQ.toLocaleString()} MT</span></div>
                  <div style={{height:8,borderRadius:4,background:"var(--border2)"}}><div style={{height:"100%",borderRadius:4,width:`${Math.min(pct,100)}%`,background:pct>=100?"var(--grn)":pct>=75?"var(--gold)":"var(--red)",transition:"width .4s"}}/></div>
                </div>;
              })}
            </div>
          </div>
          <div className="card"><div className="card-hdr"><div className="card-title">Region-wise Achievement</div></div>
            <div className="card-body">
              {[...new Set(enriched.map(r=>r.region).filter(Boolean))].map(reg=>{
                const rr=enriched.filter(r=>r.region===reg);
                const rQ=rr.reduce((s,r)=>s+Number(r.qty||0),0);
                const iQ=rr.reduce((s,r)=>s+r.totalInvQty,0);
                const pct=rQ>0?Math.round(iQ/rQ*100):0;
                return <div key={reg} style={{marginBottom:14}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}><b>{reg}</b><span style={{color:ach(pct),fontWeight:800}}>{pct}% · {iQ.toLocaleString()} / {rQ.toLocaleString()} MT</span></div>
                  <div style={{height:8,borderRadius:4,background:"var(--border2)"}}><div style={{height:"100%",borderRadius:4,width:`${Math.min(pct,100)}%`,background:pct>=100?"var(--grn)":pct>=75?"var(--gold)":"var(--red)",transition:"width .4s"}}/></div>
                </div>;
              })}
              {enriched.every(r=>!r.region)&&<div style={{color:"var(--muted)",textAlign:"center",padding:20}}>No region data.</div>}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════════
// SALES ENTRY MODAL
// ══════════════════════════════════════════════════════
// ─── Cost of Production Master ────────────────────────────
function COPMasterPanel({ copMaster, setCopMaster, plantMaster }) {
  const plants = (plantMaster||[]).map(p=>p.name);
  const COLS = [
    {key:"id",label:"ID"},{key:"plant",label:"Plant / Depot"},{key:"grade",label:"Grade"},
    {key:"materialType",label:"Material Type"},{key:"costPMT",label:"Cost PMT (₹)"},
    {key:"effectiveFrom",label:"Effective From"},{key:"effectiveTo",label:"Effective To"},{key:"remarks",label:"Remarks"}
  ];
  return (
    <InlineEditTable title="Cost of Production Master" icon="🏭"
      data={copMaster} setData={setCopMaster}
      cols={COLS} keyField="id" numFields={["costPMT"]}
      previewCols={["id","plant","grade","costPMT"]}
    />
  );
}

// ─── Packing Cost Master ──────────────────────────────────
function PackingMasterPanel({ packingMaster, setPackingMaster }) {
  const COLS = [
    {key:"id",label:"ID"},{key:"materialType",label:"Material Type"},{key:"packingType",label:"Packing Type"},
    {key:"costPerMT",label:"Cost per MT (₹)"},{key:"effectiveFrom",label:"Effective From"},
    {key:"effectiveTo",label:"Effective To"},{key:"remarks",label:"Remarks"}
  ];
  return (
    <InlineEditTable title="Packing Cost Master" icon="📦"
      data={packingMaster} setData={setPackingMaster}
      cols={COLS} keyField="id" numFields={["costPerMT"]}
      previewCols={["id","materialType","packingType","costPerMT"]}
    />
  );
}

// ─── Plant / Depot Master ─────────────────────────────────
function PlantMasterPanel({ plantMaster, setPlantMaster }) {
  const COLS = [
    {key:"id",label:"ID"},{key:"code",label:"SAP Code"},{key:"name",label:"Plant / Depot Name"},
    {key:"type",label:"Type"},{key:"state",label:"State"},{key:"region",label:"Region"},
    {key:"address",label:"Address"},{key:"capacity",label:"Capacity (MT/day)"},
    {key:"status",label:"Status"},{key:"remarks",label:"Remarks"}
  ];
  return (
    <InlineEditTable title="Plant / Depot Master" icon="🏗"
      data={plantMaster} setData={setPlantMaster}
      cols={COLS} keyField="id"
      previewCols={["code","name","type","state","status"]}
    />
  );
}

// ─── OP Commission Rate Master ────────────────────────────────
function OPCommissionMasterPanel({ opCommissionMaster, setOpCommissionMaster }) {
  const COLS = [
    {key:"id",label:"ID"},
    {key:"grade",label:"Grade"},
    {key:"paymentTerms",label:"Payment Terms"},
    {key:"commissionRate",label:"OP Commission Rate (₹/MT)"},
    {key:"effectiveFrom",label:"Effective From"},
    {key:"effectiveTo",label:"Effective To"},
    {key:"remarks",label:"Remarks"}
  ];
  return (
    <div>
      <div className="alert a-info" style={{marginBottom:14}}>
        <span>ℹ️</span>
        <div>The OP Commission Rate is auto-fetched in the price approval form based on the selected <strong>Grade</strong> and <strong>Payment Terms</strong>. Admin can override here.</div>
      </div>
      <InlineEditTable title="OP Commission Rate Master" icon="💸"
        data={opCommissionMaster} setData={setOpCommissionMaster}
        cols={COLS} keyField="id" numFields={["commissionRate"]}
        previewCols={["id","grade","paymentTerms","commissionRate"]}
      />
    </div>
  );
}

function AdminMasters({ tpcAgents, setTpcAgents, users, setUsers, customers, priceMaster, setPriceMaster, freightMaster, setFreightMaster, locationMaster, setLocationMaster, copMaster, setCopMaster, packingMaster, setPackingMaster, plantMaster, setPlantMaster, modeMaster, setModeMaster, unitSourceMaster, setUnitSourceMaster, sourceMaster, setSourceMaster, storageLocationMaster, setStorageLocationMaster, opCommissionMaster, setOpCommissionMaster }) {
  const [tab, setTab] = useState("price");
  const SIMPLE_COLS = [{key:"id",label:"ID"},{key:"name",label:"Name"}];
  return (
    <div>
      <div className="ph"><div><div className="ph-eyebrow">Administration</div><div className="ph-title">Admin Masters</div><div className="ph-sub">Manage reference data, masters and access rights</div></div></div>
      <div className="tabs">{[{ k: "price", l: "💰 Price Master" }, { k: "freight", l: "🚛 Freight Master" }, { k: "location", l: "📍 Location Master" }, { k: "cop", l: "🏭 Cost of Production" }, { k: "opcomm", l: "💸 OP Commission" }, { k: "packing", l: "📦 Packing Cost" }, { k: "plant", l: "🏗 Plant / Depot" }, { k: "agents", l: "🤝 TPC Agents" }, { k: "mode", l: "🚚 Mode" }, { k: "unitSource", l: "📐 Unit / Source" }, { k: "source", l: "🏭 Source" }, { k: "storageLocation", l: "🏪 Storage Location" }, { k: "users", l: "👤 Users" }].map(t => <div key={t.k} className={`tab ${tab === t.k ? "active" : ""}`} onClick={() => setTab(t.k)}>{t.l}</div>)}</div>
      {tab === "price"    && <PriceMasterPanel priceMaster={priceMaster} setPriceMaster={setPriceMaster} />}
      {tab === "freight"  && <FreightMasterPanel freightMaster={freightMaster} setFreightMaster={setFreightMaster} />}
      {tab === "location" && <LocationMasterPanel locationMaster={locationMaster} setLocationMaster={setLocationMaster} users={users} setUsers={setUsers} />}
      {tab === "cop"     && <COPMasterPanel     copMaster={copMaster}         setCopMaster={setCopMaster}         plantMaster={plantMaster} />}
      {tab === "opcomm"  && <OPCommissionMasterPanel opCommissionMaster={opCommissionMaster} setOpCommissionMaster={setOpCommissionMaster} />}
      {tab === "packing" && <PackingMasterPanel packingMaster={packingMaster}   setPackingMaster={setPackingMaster} />}
      {tab === "plant"   && <PlantMasterPanel   plantMaster={plantMaster}       setPlantMaster={setPlantMaster}     />}
      {tab === "agents"   && <InlineEditTable title="TPC Agent Master" icon="🤝" data={tpcAgents} setData={setTpcAgents} cols={TPC_COLS} keyField="id" previewCols={["id", "name"]} />}
      {tab === "mode"           && <InlineEditTable title="Mode Master" icon="🚚" data={modeMaster} setData={setModeMaster} cols={SIMPLE_COLS} keyField="id" previewCols={["id","name"]} />}
      {tab === "unitSource"     && <InlineEditTable title="Unit / Source Master" icon="📐" data={unitSourceMaster} setData={setUnitSourceMaster} cols={SIMPLE_COLS} keyField="id" previewCols={["id","name"]} />}
      {tab === "source"         && <InlineEditTable title="Source Master" icon="🏭" data={sourceMaster} setData={setSourceMaster} cols={SIMPLE_COLS} keyField="id" previewCols={["id","name"]} />}
      {tab === "storageLocation"&& <InlineEditTable title="Storage Location Master" icon="🏪" data={storageLocationMaster} setData={setStorageLocationMaster} cols={SIMPLE_COLS} keyField="id" previewCols={["id","name"]} />}
      {tab === "users"    && <UserMasterPanel users={users} setUsers={setUsers} customers={customers} locationMaster={locationMaster} />}
    </div>
  );
}

// ═══ PENDING ORDER REPORT ════════════════════════════════════
function PendingOrderReport({ requests, salesEntries }) {
  const { exportCsv, ExportModal, ToastEl } = useExport();
  const [fGrade, setFGrade] = useState(""); const [fRegion, setFRegion] = useState("");
  // Only approved, non-blocked, non-expired orders are "active pending" (not fully sold out)
  const enriched = requests
    .filter(r => r.status === "Approved" && !r.blocked)
    .map(r => {
      const entries = (salesEntries||[]).filter(e=>e.requestId===r.id);
      const soldQty = entries.reduce((s,e)=>s+Number(e.invoiceQty||0),0);
      const soldValue = entries.reduce((s,e)=>s+Number(e.invoiceValue||0),0);
      const approvedQty = Number(r.qty||0);
      const pendingQty = Math.max(0, approvedQty - soldQty);
      const pct = approvedQty>0 ? Math.round(soldQty/approvedQty*100) : 0;
      const isExpired = r.validityTo && r.validityTo < TODAY;
      return {...r, soldQty, soldValue, pendingQty, achievePct:pct, invoiceCount:entries.length, isExpired};
    });

  const filtered = enriched.filter(r=>{
    if (fGrade && r.grade !== fGrade) return false;
    if (fRegion && r.region !== fRegion) return false;
    return true;
  });

  const totalApproved = filtered.reduce((s,r)=>s+Number(r.qty||0),0);
  const totalSold     = filtered.reduce((s,r)=>s+r.soldQty,0);
  const totalPending  = filtered.reduce((s,r)=>s+r.pendingQty,0);
  const totalSoldVal  = filtered.reduce((s,r)=>s+r.soldValue,0);

  const REPORT_COLS = [
    {key:"id",label:"Request ID"},{key:"date",label:"Date"},{key:"customerName",label:"Customer"},
    {key:"grade",label:"Grade"},{key:"materialType",label:"Material"},{key:"region",label:"Region"},
    {key:"paymentTerms",label:"Payment Terms"},{key:"validityFrom",label:"Validity From"},{key:"validityTo",label:"Validity To"},
    {key:"qty",label:"Approved Qty (MT)"},{key:"soldQty",label:"Sold Qty (MT)"},{key:"pendingQty",label:"Pending Qty (MT)"},
    {key:"achievePct",label:"% Achieved"},{key:"soldValue",label:"Sold Value (₹)"},{key:"invoiceCount",label:"Invoices"},
  ];

  const getColor = pct => pct>=100?"var(--grn)":pct>=75?"var(--gold)":pct>=50?"var(--ora)":"var(--red)";
  const regions = [...new Set(requests.map(r=>r.region).filter(Boolean))];

  return (
    <div>
      {ToastEl}
      <div className="ph">
        <div><div className="ph-eyebrow">Reports</div><div className="ph-title">Pending Order Report</div><div className="ph-sub">Approved quantity vs actual sales — track unfulfilled orders</div></div>
        <button className="btn btn-primary btn-sm" onClick={()=>exportCsv(filtered, REPORT_COLS, `pending_order_report_${new Date().toISOString().split("T")[0]}.csv`, "Pending Order Report")}>⬇ Export</button>
      </div>
      {/* KPIs */}
      <div className="kpi-grid" style={{gridTemplateColumns:"repeat(4,1fr)",marginBottom:16}}>
        {[
          {l:"Active Orders",v:filtered.length,i:"📋",c:"var(--gold)"},
          {l:"Total Approved Qty",v:`${totalApproved.toLocaleString()} MT`,i:"📦",c:"#7C3AED"},
          {l:"Total Sold Qty",v:`${totalSold.toLocaleString()} MT`,i:"✅",c:"var(--grn)"},
          {l:"Pending (Unsold)",v:`${totalPending.toLocaleString()} MT`,i:"⏳",c:"var(--red)"},
        ].map((k,i)=>(
          <div key={i} className="kpi"><div className="kpi-accent" style={{background:k.c}}/><div className="kpi-icon">{k.i}</div>
            <div className="kpi-val" style={{color:k.c,fontSize:k.v.toString().length>8?16:24}}>{k.v}</div>
            <div className="kpi-label">{k.l}</div>
          </div>
        ))}
      </div>
      {/* Filters */}
      <div className="card" style={{marginBottom:14}}>
        <div className="card-body" style={{padding:"12px 18px"}}>
          <div style={{display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
            <select className="fsel" value={fGrade} onChange={e=>setFGrade(e.target.value)}><option value="">All Grades</option><option>OPC</option><option>PPC</option><option>ProMaxX</option></select>
            <select className="fsel" value={fRegion} onChange={e=>setFRegion(e.target.value)}><option value="">All Regions</option>{regions.map(r=><option key={r}>{r}</option>)}</select>
            {(fGrade||fRegion) && <button className="btn btn-ghost btn-sm" onClick={()=>{setFGrade("");setFRegion("");}}>✕ Clear</button>}
            <div style={{marginLeft:"auto",fontSize:13,fontWeight:700,color:"var(--grn)"}}>
              Overall Achievement: <span style={{fontFamily:"'Sora',sans-serif",fontSize:18}}>{totalApproved>0?Math.round(totalSold/totalApproved*100):0}%</span>
              &nbsp;· Sold Value: <span style={{color:"var(--gold-dk)"}}>₹{(totalSoldVal/100000).toFixed(1)}L</span>
            </div>
          </div>
        </div>
      </div>
      {/* Table */}
      <div className="card">
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12.5}}>
            <thead>
              <tr style={{background:"var(--gold-f)"}}>
                <th style={{padding:"10px 12px",textAlign:"left",fontWeight:700,borderBottom:"2px solid var(--gold-200)"}}>Request ID</th>
                <th style={{padding:"10px 12px",textAlign:"left",fontWeight:700,borderBottom:"2px solid var(--gold-200)"}}>Customer</th>
                <th style={{padding:"10px 12px",textAlign:"left",fontWeight:700,borderBottom:"2px solid var(--gold-200)"}}>Grade</th>
                <th style={{padding:"10px 12px",textAlign:"left",fontWeight:700,borderBottom:"2px solid var(--gold-200)"}}>Region</th>
                <th style={{padding:"10px 12px",textAlign:"left",fontWeight:700,borderBottom:"2px solid var(--gold-200)"}}>Validity To</th>
                <th style={{padding:"10px 12px",textAlign:"right",fontWeight:700,borderBottom:"2px solid var(--gold-200)"}}>Approved Qty (MT)</th>
                <th style={{padding:"10px 12px",textAlign:"right",fontWeight:700,borderBottom:"2px solid var(--gold-200)"}}>Sold Qty (MT)</th>
                <th style={{padding:"10px 12px",textAlign:"right",fontWeight:700,borderBottom:"2px solid var(--gold-200)"}}>Pending Qty (MT)</th>
                <th style={{padding:"10px 12px",textAlign:"right",fontWeight:700,borderBottom:"2px solid var(--gold-200)"}}>% Achieved</th>
                <th style={{padding:"10px 12px",textAlign:"right",fontWeight:700,borderBottom:"2px solid var(--gold-200)"}}>Sold Value (₹)</th>
                <th style={{padding:"10px 12px",textAlign:"right",fontWeight:700,borderBottom:"2px solid var(--gold-200)"}}>Invoices</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length===0 && <tr><td colSpan={11} style={{textAlign:"center",padding:40,color:"var(--muted)"}}>No approved orders found.</td></tr>}
              {filtered.map((r,i)=>(
                <tr key={r.id} style={{background:i%2===0?"var(--white)":r.isExpired?"#FFF0E0":"var(--cream)"}}>
                  <td style={{padding:"8px 12px",fontFamily:"'Sora',sans-serif",fontWeight:800,color:"var(--gold-dk)",fontSize:12}}>{r.id}{r.isExpired&&<span className="badge b-completed" style={{marginLeft:4,fontSize:10}}>Expired</span>}</td>
                  <td style={{padding:"8px 12px",fontWeight:600}}>{r.customerName}</td>
                  <td style={{padding:"8px 12px"}}><span className="badge b-grade">{r.grade}</span></td>
                  <td style={{padding:"8px 12px",color:"var(--muted)"}}>{r.region}</td>
                  <td style={{padding:"8px 12px",color:r.isExpired?"var(--red)":"var(--muted)",fontWeight:r.isExpired?700:400}}>{r.validityTo}</td>
                  <td style={{padding:"8px 12px",textAlign:"right",fontWeight:700}}>{Number(r.qty).toLocaleString()}</td>
                  <td style={{padding:"8px 12px",textAlign:"right",fontWeight:700,color:"var(--grn)"}}>{r.soldQty.toLocaleString()}</td>
                  <td style={{padding:"8px 12px",textAlign:"right",fontWeight:800,color:r.pendingQty>0?"var(--red)":"var(--grn)"}}>{r.pendingQty.toLocaleString()}</td>
                  <td style={{padding:"8px 12px",textAlign:"right"}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,justifyContent:"flex-end"}}>
                      <div style={{width:60,height:6,borderRadius:4,background:"var(--border2)"}}>
                        <div style={{height:"100%",borderRadius:4,width:`${Math.min(r.achievePct,100)}%`,background:getColor(r.achievePct)}}/>
                      </div>
                      <span style={{fontWeight:800,color:getColor(r.achievePct),fontFamily:"'Sora',sans-serif"}}>{r.achievePct}%</span>
                    </div>
                  </td>
                  <td style={{padding:"8px 12px",textAlign:"right",color:"var(--grn)",fontWeight:700}}>₹{(r.soldValue/1000).toFixed(0)}K</td>
                  <td style={{padding:"8px 12px",textAlign:"right",color:"var(--muted)"}}>{r.invoiceCount}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
      {ExportModal}
    </div>
  );
}

// ═══ REPORTS PAGE ════════════════════════════════════════
function ReportsPage({ requests, salesEntries }) {
  const [reportTab, setReportTab] = useState("analytics");
  const [dateFrom, setDateFrom] = useState(""); const [dateTo, setDateTo] = useState(""); const [fStatus, setFStatus] = useState(""); const [fGrade, setFGrade] = useState(""); const [fRegion, setFRegion] = useState("");
  const { exportCsv, exportTemplate, ExportModal, ToastEl } = useExport();
  const filtered = requests.filter(r => { if (fStatus && r.status !== fStatus) return false; if (fGrade && r.grade !== fGrade) return false; if (fRegion && r.region !== fRegion) return false; if (dateFrom && r.date < dateFrom) return false; if (dateTo && r.date > dateTo) return false; return true; });
  const totalQty = filtered.reduce((s, r) => s + Number(r.qty || 0), 0);
  const avgNcr = filtered.length ? filtered.reduce((s, r) => s + Number(r.ncrPmt || 0), 0) / filtered.length : 0;
  const totalOrderVal = filtered.reduce((s, r) => s + Number(r.orderPrice || 0) * Number(r.qty || 0), 0);
  const regions = [...new Set(requests.map(r => r.region).filter(Boolean))];
  const gradeStats = ["OPC", "PPC", "ProMaxX"].map(g => { const gr = filtered.filter(r => r.grade === g); const avgN = gr.length ? gr.reduce((s, r) => s + Number(r.ncrPmt || 0), 0) / gr.length : 0; const vol = gr.reduce((s, r) => s + Number(r.qty || 0), 0); return { g, count: gr.length, vol, avgNcr: avgN.toFixed(0), approved: gr.filter(r => r.status === "Approved").length }; });
  const ptStats = [...new Set(filtered.map(r => r.paymentTerms))].map(pt => ({ pt, n: filtered.filter(r => r.paymentTerms === pt).length })).sort((a, b) => b.n - a.n);
  const maxPt = Math.max(...ptStats.map(p => p.n), 1);
  const custVol = [...new Set(filtered.map(r => r.customerCode))].map(c => { const cr = filtered.filter(r => r.customerCode === c); return { code: c, name: cr[0]?.customerName || c, vol: cr.reduce((s, r) => s + Number(r.qty || 0), 0), count: cr.length }; }).sort((a, b) => b.vol - a.vol).slice(0, 5);
  return (
    <div>
      {ToastEl}
      <div className="ph">
        <div><div className="ph-eyebrow">Analytics</div><div className="ph-title">Reports</div><div className="ph-sub">Analyse and export price request data</div></div>
        {reportTab==="analytics" && <div className="ph-actions">
          <button className="btn btn-secondary btn-sm" onClick={() => exportTemplate(REQUEST_REPORT_COLS, "report_template.csv")}>📋 Template</button>
          <button className="btn btn-primary btn-sm" onClick={() => exportCsv(filtered, REQUEST_REPORT_COLS, `report_${new Date().toISOString().split("T")[0]}.csv`, `Report (${filtered.length} records)`)}>⬇ Export CSV</button>
        </div>}
      </div>
      <div className="tabs" style={{marginBottom:20}}>
        <div className={`tab ${reportTab==="analytics"?"active":""}`} onClick={()=>setReportTab("analytics")}>📊 Analytics Report</div>
        <div className={`tab ${reportTab==="pending"?"active":""}`} onClick={()=>setReportTab("pending")}>⏳ Pending Order Report</div>
      </div>
      {reportTab==="pending" && <PendingOrderReport requests={requests} salesEntries={salesEntries} />}
      {reportTab==="analytics" && <>
      <div className="card" style={{ marginBottom: 20 }}><div className="card-hdr"><div className="card-title">🔍 Filter Report</div></div>
        <div className="card-body">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(160px,1fr))", gap: 14 }}>
            <div className="fg"><label className="fl">Date From</label><input className="fc" type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)} /></div>
            <div className="fg"><label className="fl">Date To</label><input className="fc" type="date" value={dateTo} onChange={e => setDateTo(e.target.value)} /></div>
            <div className="fg"><label className="fl">Status</label><select className="fc" value={fStatus} onChange={e => setFStatus(e.target.value)}><option value="">All</option><option>Pending</option><option>Approved</option><option>Rejected</option></select></div>
            <div className="fg"><label className="fl">Grade</label><select className="fc" value={fGrade} onChange={e => setFGrade(e.target.value)}><option value="">All</option><option>OPC</option><option>PPC</option><option>ProMaxX</option></select></div>
            <div className="fg"><label className="fl">Region</label><select className="fc" value={fRegion} onChange={e => setFRegion(e.target.value)}><option value="">All</option>{regions.map(r => <option key={r}>{r}</option>)}</select></div>
          </div>
          {(dateFrom || dateTo || fStatus || fGrade || fRegion) && <button className="btn btn-ghost btn-sm" style={{ marginTop: 12 }} onClick={() => { setDateFrom(""); setDateTo(""); setFStatus(""); setFGrade(""); setFRegion(""); }}>✕ Clear</button>}
        </div>
      </div>
      <div className="kpi-grid">
        {[{ l: "Filtered", v: filtered.length, i: "📋", c: "var(--gold)", bg: "var(--gold-f)" }, { l: "Volume (MT)", v: totalQty.toLocaleString(), i: "📦", c: "#7C3AED", bg: "#F5F3FF" }, { l: "Order Value", v: `₹${(totalOrderVal / 1000).toFixed(0)}K`, i: "💰", c: "var(--grn)", bg: "var(--grn-f)" }, { l: "Avg NCR/MT", v: `₹${avgNcr.toFixed(0)}`, i: "📊", c: avgNcr < -300 ? "var(--red)" : "var(--grn)", bg: avgNcr < -300 ? "var(--red-f)" : "var(--grn-f)" }, { l: "Approval Rate", v: `${filtered.length ? Math.round(filtered.filter(r => r.status === "Approved").length / filtered.length * 100) : 0}%`, i: "✅", c: "var(--grn)", bg: "var(--grn-f)" }].map((k, i) => (
          <div key={i} className="kpi"><div className="kpi-accent" style={{ background: k.c }} /><div className="kpi-icon" style={{ background: k.bg }}>{k.i}</div><div className="kpi-val" style={{ fontSize: k.v.toString().length > 6 ? 20 : 28 }}>{k.v}</div><div className="kpi-label">{k.l}</div></div>
        ))}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
        <div className="card">
          <div className="card-hdr"><div className="card-title">Grade-wise Analysis</div></div>
          <div className="card-body">
            {gradeStats.map(g => (
              <div key={g.g} style={{ padding: "12px 0", borderBottom: "1px solid var(--border2)" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                  <span className="badge b-grade">{g.g}</span>
                  <span style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 15, color: Number(g.avgNcr) < -300 ? "var(--red)" : "var(--grn)" }}>NCR: ₹{g.avgNcr}</span>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
                  {[{ l: "Requests", v: g.count }, { l: "Volume (MT)", v: g.vol.toLocaleString() }, { l: "Approved", v: g.approved }].map(s => (
                    <div key={s.l} style={{ background: "var(--cream)", borderRadius: 8, padding: "8px 10px", textAlign: "center" }}>
                      <div style={{ fontFamily: "'Sora',sans-serif", fontSize: 16, fontWeight: 800, color: "var(--ink)" }}>{s.v}</div>
                      <div style={{ fontSize: 10.5, fontWeight: 700, color: "var(--muted)", textTransform: "uppercase", letterSpacing: .4 }}>{s.l}</div>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
          <div className="card">
            <div className="card-hdr"><div className="card-title">Payment Terms</div></div>
            <div className="card-body">
              <div className="bar-chart">
                {ptStats.map((p, i) => (
                  <div key={p.pt} className="bar-row">
                    <div className="bar-label">{p.pt}</div>
                    <div className="bar-track"><div className="bar-fill" style={{ width: `${(p.n / maxPt) * 100}%`, background: ["var(--gold)", "#0EA5E9", "#8B5CF6", "#10B981"][i % 4] }}>{p.n > 0 ? p.n : ""}</div></div>
                    <div className="bar-val">{p.n}</div>
                  </div>
                ))}
                {!ptStats.length && <div style={{ textAlign: "center", color: "var(--muted)", padding: 20 }}>No data</div>}
              </div>
            </div>
          </div>
          <div className="card">
            <div className="card-hdr"><div className="card-title">Top Customers by Volume</div></div>
            <div className="card-body" style={{ padding: "8px 20px" }}>
              {custVol.map((c, i) => (
                <div key={c.code} style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 0", borderBottom: i < custVol.length - 1 ? "1px solid var(--border2)" : "none" }}>
                  <div style={{ width: 28, height: 28, borderRadius: 8, background: "var(--gold-f)", border: "1.5px solid var(--gold-200)", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 13, color: "var(--gold-dk)", flexShrink: 0 }}>{i + 1}</div>
                  <div style={{ flex: 1, minWidth: 0 }}><div style={{ fontWeight: 700, fontSize: 13, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{c.name}</div><div style={{ fontSize: 11.5, color: "var(--muted)" }}>{c.count} request{c.count > 1 ? "s" : ""}</div></div>
                  <div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 14, color: "var(--gold-dk)" }}>{c.vol.toLocaleString()} MT</div>
                </div>
              ))}
              {!custVol.length && <div style={{ textAlign: "center", color: "var(--muted)", padding: 20 }}>No data</div>}
            </div>
          </div>
        </div>
      </div>
      <div className="card">
        <div className="card-hdr"><div><div className="card-title">Detailed Report</div><div className="card-sub">{filtered.length} records</div></div>
          <button className="btn btn-secondary btn-sm" onClick={() => exportCsv(filtered, REQUEST_REPORT_COLS, `report_${new Date().toISOString().split("T")[0]}.csv`, `Report (${filtered.length})`)}>⬇ Export</button>
        </div>
        <div className="tw"><table>
          <thead><tr><th>Request ID</th><th>Date</th><th>Customer</th><th>Grade</th><th>Qty</th><th>Order Price</th><th>NCR/MT</th><th>Status</th><th>Level</th></tr></thead>
          <tbody>
            {filtered.length === 0 && <tr><td colSpan={9} style={{ textAlign: "center", padding: "40px", color: "var(--muted)" }}>No records match filters</td></tr>}
            {filtered.map(r => <tr key={r.id}><td style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: "var(--gold-dk)", fontSize: 12.5 }}>{r.id}</td><td style={{ color: "var(--muted)", fontSize: 12 }}>{r.date}</td><td><div style={{ fontWeight: 700 }}>{r.customerName}</div><div style={{ fontSize: 11, color: "var(--muted)" }}>{r.region}</div></td><td><span className="badge b-grade">{r.grade}/{r.materialType}</span></td><td style={{ fontWeight: 700 }}>{r.qty}</td><td style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800 }}>₹{r.orderPrice}</td><td style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: r.ncrPmt < -300 ? "var(--red)" : "var(--grn)" }}>₹{Number(r.ncrPmt).toFixed(0)}</td><td><StatusBadge status={r.status} /></td><td style={{ fontSize: 12, color: "var(--muted)" }}>{r.currentLevel}</td></tr>)}
          </tbody>
        </table></div>
      </div>
      {ExportModal}
      </>}
    </div>
  );
}

// ═══ AUDIT LOG ════════════════════════════════════════════
function AuditLog({ requests, users }) {
  const AUDIT_COLS = [{ key: "reqId", label: "Request ID" }, { key: "level", label: "Level" }, { key: "action", label: "Action" }, { key: "by", label: "User ID" }, { key: "time", label: "Timestamp" }, { key: "remark", label: "Remark" }];
  const logs = requests.flatMap(r => r.history.map((h, i) => ({ ...h, reqId: r.id, key: `${r.id}-${i}` }))).reverse();
  const { exportCsv, exportTemplate, ExportModal, ToastEl } = useExport();
  return (
    <div>
      {ToastEl}
      <div className="ph">
        <div><div className="ph-eyebrow">Compliance</div><div className="ph-title">Audit Log</div><div className="ph-sub">{logs.length} entries</div></div>
        <div className="ph-actions">
          <button className="btn btn-secondary btn-sm" onClick={() => exportTemplate(AUDIT_COLS, "audit_template.csv")}>📋 Template</button>
          <button className="btn btn-secondary btn-sm" onClick={() => exportCsv(logs, AUDIT_COLS, `audit_log_${new Date().toISOString().split("T")[0]}.csv`, `Audit Log (${logs.length})`)}>⬇ Export</button>
        </div>
      </div>
      <div className="card"><div className="tw"><table>
        <thead><tr><th>Request ID</th><th>Action</th><th>By</th><th>Level</th><th>Timestamp</th><th>Remark</th></tr></thead>
        <tbody>{logs.map(h => (
          <tr key={h.key}>
            <td style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: "var(--gold-dk)", fontSize: 12.5 }}>{h.reqId}</td>
            <td><StatusBadge status={h.action === "Approved" || h.action === "Validated" ? "Approved" : h.action === "Rejected" ? "Rejected" : "Validated"} /></td>
            <td><div style={{ fontWeight: 700, fontSize: 13 }}>{users.find(u => u.id === h.by)?.name || h.by}</div><div style={{ fontSize: 11, color: "var(--muted)" }}>{h.by}</div></td>
            <td style={{ fontSize: 12.5, fontWeight: 600 }}>{h.level}</td>
            <td style={{ fontSize: 12, color: "var(--muted)" }}>{h.time}</td>
            <td style={{ fontSize: 12.5, fontStyle: h.remark ? "normal" : "italic", color: h.remark ? "var(--ink)" : "var(--muted)" }}>{h.remark || "—"}</td>
          </tr>
        ))}</tbody>
      </table></div></div>
      {ExportModal}
    </div>
  );
}

// ═══ NCR CALCULATOR ════════════════════════════════════════
const NCR_GST = 0.18;
const NCR_MT_BAGS = 20;

const NCR_C = {
  navy:"#0D1B2E", navyMid:"#132340", navyLt:"#1A2E50",
  gold:"#D4890A", goldLt:"#F5B800", goldF:"#FFF8EC",
  cream:"#FFFDF5", white:"#FFFFFF",
  ink:"#1A1200", ink3:"#6B5200",
  muted:"#9C8050", border:"#EDE0C0", border2:"#F5EDD5",
  grn:"#059669", grnF:"#ECFDF5", grnBd:"#A7F3D0",
  red:"#DC2626", redF:"#FEF2F2",
  mangRed:"#C0142E", blue:"#2563EB",
};

const NCR_FILTER_LABELS = {
  salesOffice:"Sales Office",
  shipZone:"Zone",
  soldToCode:"Sold-to-party Code",
  soldToName:"Sold-to-party Name",
  grade:"Grade (PPC / OPC / PMX / ALL)",
  shipCity:"Ship City",
  priceBlock:"SHP Sales District / Price Block / Taluk",
  mode:"Mode",
  unitSource:"Unit / Source",
  source:"Source",
  storageLocation:"Storage Location",
};

const NCR_COST_KEYS = [
  ["cashDiscountGST",  "Cash Discount (NET OF GST)"],
  ["cashDiscountDN",   "Cash Discount (DEBIT NOTE)"],
  ["incentiveQDM",     "INCENTIVE (QD-M) a"],
  ["incentiveQDQRT",   "INCENTIVE (QD-QRT) b"],
  ["incentiveQDH",     "INCENTIVE (QD-H) c"],
  ["incentiveSTS",     "INCENTIVE (STS)"],
  ["retailer1",        "RETAILER-1 (GENERAL)"],
  ["retailer2",        "RETAILER-2 (STS)"],
  ["opComm",           "OP COMMISSION"],
  ["spComm",           "SP COMMISSION"],
  ["moComm",           "MO COMMISSION"],
  ["distComm",         "Distribution Commission"],
  ["secFreight",       "Secondary Freight"],
  ["unloading",        "Unloading Charges"],
  ["primFreight",      "Primary Freight"],
  ["trolla",           "TROLLA"],
  ["godownShifting",   "Godown Shifting"],
  ["godownHandling",   "Godown Handling"],
  ["stationHandling",  "Station Handling"],
  ["demurrage",        "Demurrage & Wharfage"],
  ["packing",          "Packing"],
];

const NCR_DEFAULT_COSTS = {
  cashDiscountGST:  { prev:49.504257,   curr:49.504257   },
  cashDiscountDN:   { prev:0,           curr:0           },
  incentiveQDM:     { prev:38.956375,   curr:38.956375   },
  incentiveQDQRT:   { prev:29.457564,   curr:29.457564   },
  incentiveQDH:     { prev:64.302954,   curr:64.302954   },
  incentiveSTS:     { prev:202.973444,  curr:202.973444  },
  retailer1:        { prev:29.530519,   curr:29.530519   },
  retailer2:        { prev:34.276721,   curr:34.276721   },
  opComm:           { prev:10.102150,   curr:10.102150   },
  spComm:           { prev:30.046350,   curr:30.046350   },
  moComm:           { prev:49.698643,   curr:49.698643   },
  distComm:         { prev:2.518485,    curr:2.518485    },
  secFreight:       { prev:169.733352,  curr:169.733352  },
  unloading:        { prev:2.766525,    curr:2.766525    },
  primFreight:      { prev:763.611663,  curr:763.611663  },
  trolla:           { prev:18.243904,   curr:18.243904   },
  godownShifting:   { prev:23.868916,   curr:23.868916   },
  godownHandling:   { prev:28.771376,   curr:28.771376   },
  stationHandling:  { prev:37.134693,   curr:37.134693   },
  demurrage:        { prev:11.742970,   curr:11.742970   },
  packing:          { prev:162.556152,  curr:162.556152  },
  cop:              { prev:3690.900128, curr:3690.900128 },
};

const NCR_EMPTY_FILTERS = Object.fromEntries(Object.keys(NCR_FILTER_LABELS).map(k=>[k,"ALL"]));

function NCRCalculator({ user }:{ user:any }) {
  const [rows, setRows]         = useState([]);
  const [fileName, setFileName] = useState("");
  // Auto-apply Sales Office restriction for non-admin users
  const userSONames = useMemo(() => {
    if (!user || user.role==="Admin") return [];
    return (user.allowedSalesOffices||[]).length > 0
      ? (user.allowedSalesOffices||[]).map((soId:string) => soId) // ids; names resolved at filter time
      : [];
  }, [user]);
  const [costs, setCosts]       = useState(NCR_DEFAULT_COSTS);
  const [filters, setFilters]   = useState(NCR_EMPTY_FILTERS);
  const [currPx, setCurrPx]     = useState("");
  const [editC, setEditC]       = useState(false);

  const fileRef = useRef();

  const handleFile = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type:"array" });
        // Prefer "COST SHEET" tab; fallback to first sheet
        const KEY_SHEETS = ["COST SHEET","Billing Data","Sheet1"];
        const sheetName = KEY_SHEETS.find(s => wb.SheetNames.includes(s)) || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        // Auto-detect header row: find first row that has known column headers
        const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
        const HDR_MARKERS = ["GRADE","BillingQuantity","TR & NT","REALSN AMOUNT","Sold-topartyCode"];
        let headerRow = 0;
        for (let i = 0; i < Math.min(8, raw.length); i++) {
          if (raw[i].some(v => HDR_MARKERS.includes(String(v).trim()))) { headerRow = i; break; }
        }
        const parsed = XLSX.utils.sheet_to_json(ws, { defval:"", range: headerRow });
        setRows(parsed);
        setFileName(file.name + "  [" + sheetName + ", " + parsed.length + " rows]");
        setFilters(NCR_EMPTY_FILTERS);
      } catch(err) { alert("Error reading file: " + err.message); }
    };
    reader.readAsArrayBuffer(file); e.target.value="";
  };

  const dlTemplate = () => {
    const hdr = [
      "GRADE","MODE","UNIT","DIRECT DEPOT","CODE","InvoiceNumber","ODN","OrderNo","OrderDate","OrderQty",
      "QuatationNo","BillingDate","MONTH","QTR","DeliveryNumber","SalesOrganization","DistributionChannel",
      "DistributionChannel","BillingType","SoldToPartyState","SoldtoPartyBranch","ShiptoPartyState",
      "ShipToPartyBranch","Ship-to-partyBranchdesc","SHIP TO DISTRICT CODE","DISTRICTNAME","STATE",
      "SALES OFFICE","Zone -1","Sub zone-2","Ship to district -3","ShiptoParty","Destination","Route",
      "Incotermspart1","Sold-topartyCode","Sold-topartyNAME","Grade/Material","ShipmentNumber","BatchNumber",
      "Source","Storagelocation","TypeOfPacking","GrossBillingValue(EX_CD&TD)","TradeDiscount-INR",
      "NetValueofBillingExcludingtaxes&CD/TD-INR","TCS","PRICE","PRICE CALC","GST ON CD-TD-LR",
      "BillingQuantity","GROSS BILLING","TotalGST","Cash Discount (NET OF GST)3","Cash Discount (DEBIT NOTE)",
      "INCENTIVE(QD-M)a","INCENTIVE(QD-QRT)b","INCENTIVE(QD-H)c","RETAILER-1 (GENERAL)","RETAILER-2 (STS)",
      "INCENTIVE(STS)","TROLLA","OPCOMISSION","SPCOMISSION","MO COMM","SecondaryFreight","UnloadingCharges",
      "Primary Freight","GodownShifting","GodownHandling","StationHandling","Dem&Wrfage","DISTRIBUTION COMM .",
      "Packing","TotalCostOfProduction","REALSN AMOUNT","NCR/MT","USP","UDAAN","Current CP","OLD CP","Diff",
      "Revised RealISATION","Revised NCR/PMT","Price","CD.","Mon.","Qtrly","Hlfy","STS..","Trolla","RTL",
      "SPC.","OPC.","MO","Dis C.","USP Rs","AK/ Rs","Discount","Actual Disc","All Disc & com","Total Logistic",
      "Prm Frt","Sec Frt","Ship ST","Ship Zone","Ship Sub Zone","Sale Office wise name","Ship State",
      "Ship District","SHP Sales District / Price Block / Taluk","Transportation Zone / Destination Desc",
      "Ship City","Sold-to party","Name","Ship-to party","Name2","INDENTING AGENT","NAME OF INDENTING AGENT",
      "Material Number","Material Description","Material Description2","Grade","TR & NT","Distribution Channel",
      "Final Month","New Grade","Grade","Quarter","New source"
    ];
    // Sample row: fill key calc columns, leave rest blank ("")
    const blank = (n:number) => Array(n).fill("");
    const row1 = [
      "PPC","Road","Aligarh Plant","Retailer","",                   // 0-4  GRADE…CODE
      "INV001","","ORD001","01-Apr-26",500,                          // 5-9
      "","01-Apr-26","Apr-26","Q1","",                               // 10-14
      "1000","Trade","Trade","","",                                  // 15-19
      "","","","","",                                                // 20-24
      "","RJ","SO-JAI","Zone A","",                                  // 25-29
      "","","JAI001","Jaipur","",                                    // 30-34
      "","C001","Raj Construction Co.","PPC001","","",               // 35-40  Sold-to…BatchNo
      "Plant","SL001","Bag",44000,2000,                              // 41-45
      42000,0,350,350,54,                                            // 46-50  PRICE…GST
      140,49000,8820,200,0,                                          // 51-56
      0,0,0,500,300,                                                 // 57-61
      0,100,50,30,80,                                                // 62-66
      20,10,5,15,25,                                                 // 67-71
      0,3600,135,-50,                                                // 72-75  COP…NCR/MT
      "","",0,350,0,                                                 // 76-80
      0,0,0,0,0,                                                     // 81-85
      0,0,0,0,0,                                                     // 86-90
      0,0,0,0,0,                                                     // 91-95
      0,0,0,0,0,                                                     // 96-100
      0,0,"RJ","Zone A","",                                          // 101-105
      "Jaipur SO","Rajasthan","Jaipur",                              // 106-108
      "AJMER","Jaipur-North","Jaipur",                               // 109-111
      "C001","Raj Construction Co.","SHP001","Raj Godown",           // 112-115
      "","","MAT001","PPC 53 Grade","PPC 53",                        // 116-120
      "PPC","Trade","Trade","Apr-26","PPC","PPC","Q1","Plant"        // 121-128 → 129 items
    ];
    const row2 = [
      "OPC","Rail","Morak Plant","Direct","",
      "INV002","","ORD002","01-Apr-26",300,
      "","01-Apr-26","Apr-26","Q1","",
      "1000","Non-Trade","Non-Trade","","",
      "","","","","",
      "","RJ","SO-KOT","Zone B","",
      "","","KOT001","Kota","",
      "","C002","Sharma Builders","OPC001","","",
      "Plant","SL002","Bag",37800,1800,
      36000,0,420,420,65,
      80,33600,6048,0,0,
      0,0,0,0,0,
      0,0,0,0,100,
      20,10,0,0,200,
      0,3600,220,30,
      "","",0,420,0,
      0,0,0,0,0,
      0,0,0,0,0,
      0,0,0,0,0,
      0,0,0,0,0,
      0,0,"RJ","Zone B","",
      "Kota SO","Rajasthan","Kota",
      "KOTA","Kota-Central","Kota",
      "C002","Sharma Builders","SHP002","Sharma Godown",
      "","","MAT002","OPC 53 Grade","OPC 53",
      "OPC","Non-Trade","Non-Trade","Apr-26","OPC","OPC","Q1","Plant"
    ];
    const ws = XLSX.utils.aoa_to_sheet([hdr, row1, row2]);
    // Set column widths hint via !cols
    ws["!cols"] = hdr.map(() => ({ wch: 16 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Billing Data");
    XLSX.writeFile(wb, "NCR_Calculator_Template.xlsx");
  };

  // Column map: filterKey → data column name
  const FILTER_COL = {
    salesOffice:    "SALES OFFICE",
    shipZone:       "Zone -1",
    soldToCode:     "Sold-topartyCode",
    soldToName:     "Sold-topartyNAME",
    grade:          "Grade",
    shipCity:       "Ship City",
    priceBlock:     "SHP Sales District / Price Block / Taluk",
    mode:           "MODE",
    unitSource:     "UNIT",
    source:         "Source",
    storageLocation:"Storagelocation",
  };

  // Cascading opts: each dropdown shows only values present when all OTHER filters active
  const opts = useMemo(() => {
    const u = (selfKey:string) => {
      const subset = rows.filter(r =>
        Object.entries(FILTER_COL).every(([k, col]) => {
          if (k === selfKey) return true;          // skip self
          const v = filters[k];
          if (!v || v === "ALL") return true;
          return String(r[col] ?? "") === v;
        })
      );
      const col = FILTER_COL[selfKey];
      const vals = [...new Set(subset.map(r=>String(r[col]??"")))].filter(v=>v.trim()!=="").sort();
      return ["ALL",...vals];
    };
    return Object.fromEntries(Object.keys(FILTER_COL).map(k=>[k, u(k)]));
  }, [rows, filters]);

  const filtered = useMemo(() => rows.filter(r => {
    const chk = (key,col) => filters[key]==="ALL" || String(r[col])===filters[key];
    // User-level Sales Office restriction (non-admin with specific SO rights)
    if (userSONames.length>0 && !userSONames.includes(String(r["SALES OFFICE"]||""))) return false;
    return chk("salesOffice","SALES OFFICE") && chk("shipZone","Zone -1") &&
           chk("soldToCode","Sold-topartyCode") && chk("soldToName","Sold-topartyNAME") &&
           chk("grade","Grade") && chk("shipCity","Ship City") &&
           chk("priceBlock","SHP Sales District / Price Block / Taluk") && chk("mode","MODE") &&
           chk("unitSource","UNIT") && chk("source","Source") &&
           chk("storageLocation","Storagelocation");
  }), [rows, filters, userSONames]);

  const agg = useMemo(() => {
    const ag = (sub) => {
      // BillingQuantity = MT; REALSN AMOUNT = total NCR ₹ (not per-MT); PRICE = ₹/bag
      const qty  = sub.reduce((s,r)=>s+(parseFloat(r["BillingQuantity"])||0),0);
      // Weighted NCR/MT = sum(REALSN AMOUNT) / sum(BillingQuantity)
      const totalReal = sub.reduce((s,r)=>s+(parseFloat(r["REALSN AMOUNT"])||0),0);
      const real = qty>0 ? totalReal/qty : 0;
      // Weighted price/bag = sum(PRICE * BillingQty) / sum(BillingQty)
      const price= qty>0 ? sub.reduce((s,r)=>s+(parseFloat(r["PRICE"])||0)*(parseFloat(r["BillingQuantity"])||0),0)/qty : 0;
      return { qty:Math.round(qty*10)/10, realization:Math.round(real*100)/100, price:Math.round(price) };
    };
    // DistributionChannel: "TR"=Trade, "NT"=Non-Trade, "NCM"=NCM
    const bt  = (t) => filtered.filter(r=>String(r["DistributionChannel"]||"").toUpperCase()===t.toUpperCase());
    const trRows = bt("TR");
    const trade= ag(trRows), nt=ag(bt("NT")), ncm=ag(bt("NCM"));
    // Direct% and Retailer% explicitly on TR billing quantity only
    const trDirQ  = trRows.filter(r=>String(r["DIRECT DEPOT"]||"").toUpperCase()==="DIRECT").reduce((s,r)=>s+(parseFloat(r["BillingQuantity"])||0),0);
    // Retailer% = CODE starting with "93" in TR rows / TR total qty (matches Excel: 65.71%)
    const trRetQ  = trRows.filter(r=>String(r["CODE"]??"").startsWith("93")).reduce((s,r)=>s+(parseFloat(r["BillingQuantity"])||0),0);
    const trQty   = trade.qty;
    const directPct   = trQty>0 ? Math.round(trDirQ/trQty*10000)/100 : 0;   // 2-decimal %
    const retailerPct = trQty>0 ? Math.round(trRetQ/trQty*10000)/100 : 0;
    const aQ   = trade.qty+nt.qty+ncm.qty;
    // Overall weighted avg NCR/MT = sum of totalReal / total qty
    const allReal = filtered.reduce((s,r)=>s+(parseFloat(r["REALSN AMOUNT"])||0),0);
    const ovR  = aQ>0 ? allReal/aQ : 0;
    const ovP  = aQ>0?(trade.price*trade.qty+nt.price*nt.qty+ncm.price*ncm.qty)/aQ:0;
    // Overall Price = PRICE CALC weighted avg (matches Excel formula)
    const ovPriceCalcWt = filtered.reduce((s,r)=>s+(parseFloat(r["PRICE CALC"])||0)*(parseFloat(r["BillingQuantity"])||0),0);
    const ovPriceCalc   = aQ>0 ? ovPriceCalcWt/aQ : 0;
    // Overall GST = TotalGST / (BillingQty × 20 bags)  (matches Excel TotalGST/bags formula)
    const ovTotalGST    = filtered.reduce((s,r)=>s+(parseFloat(r["TotalGST"])||0),0);
    const ovGSTperBag   = aQ>0 ? ovTotalGST/(aQ*20) : 0;
    const months= [...new Set(rows.map(r=>String(r["MONTH"]||"")).filter(Boolean))].sort();
    const tradeWithPct = {...trade, directPct, retailerPct};
    return { trade:tradeWithPct, nt, ncm,
      overall:{ qty:aQ, realization:Math.round(ovR*100)/100, price:Math.round(ovP) },
      prevPrice:Math.round(ovPriceCalc*100)/100,
      prevGST:  Math.round(ovGSTperBag*100)/100,
      prevNet:  Math.round((ovPriceCalc-ovGSTperBag)*100)/100,
      prevNCR:  Math.round(ovR*100)/100,
      lastMonth:months.length?months[months.length-1]:"—" };
  }, [filtered, rows]);

  const prevTotalExp = NCR_COST_KEYS.reduce((s,[k])=>s+costs[k].prev,0);
  const currTotalExp = NCR_COST_KEYS.reduce((s,[k])=>s+costs[k].curr,0);

  const prevCalc = useMemo(()=>({
    price: agg.prevPrice,
    gst:   agg.prevGST,
    net:   agg.prevNet,
    ncr:   agg.prevNCR,
  }), [agg]);

  const currCalc = useMemo(()=>{
    const price = parseFloat(currPx)||0;
    if (!price) return { price:0, gst:0, net:0, ncr:null };
    const gst = Math.round(price*NCR_GST/(1+NCR_GST));
    const net = Math.round(price/(1+NCR_GST)*100)/100;
    const ncr = Math.round(net*NCR_MT_BAGS - currTotalExp - costs.cop.curr);
    return { price, gst, net, ncr };
  }, [currPx, currTotalExp, costs.cop.curr]);

  const noData = rows.length===0;
  const setF = (k:string, v:string) => {
    setFilters(prev => {
      const next = {...prev, [k]: v};
      // Reset any filter whose current value is no longer valid after this change
      Object.keys(FILTER_COL).forEach(fk => {
        if (fk === k) return;
        if (next[fk] === "ALL") return;
        // Check if current value still exists in new filtered subset
        const col = FILTER_COL[fk];
        const stillValid = rows.some(r =>
          Object.entries(FILTER_COL).every(([ck, cc]) => {
            const fv = ck === fk ? next[fk] : next[ck];
            if (!fv || fv === "ALL") return true;
            if (ck === fk) return String(r[cc] ?? "") === fv;
            return String(r[cc] ?? "") === fv;
          })
        );
        if (!stillValid) next[fk] = "ALL";
      });
      return next;
    });
  };
  const setCost= (key,period,val) => setCosts(c=>({...c,[key]:{...c[key],[period]:parseFloat(val)||0}}));

  const C = NCR_C;
  const tdL = { padding:"7px 14px", fontSize:13, fontWeight:600, color:C.ink, borderBottom:`1px solid ${C.border2}`, verticalAlign:"middle" };
  const tdV = { padding:"7px 14px", textAlign:"right", fontWeight:700, color:C.ink, borderBottom:`1px solid ${C.border2}`, verticalAlign:"middle" };
  const tdN = { padding:"7px 8px", fontSize:11, color:C.muted, fontStyle:"italic", textAlign:"right", whiteSpace:"nowrap" as const, borderBottom:`1px solid ${C.border2}`, verticalAlign:"middle" };

  const NCard = ({children,style={}}:{children:any,style?:any}) => (
    <div style={{ background:C.white, border:`1.5px solid ${C.border}`, borderRadius:12, overflow:"hidden", boxShadow:"0 2px 8px rgba(180,130,0,.09)", ...style }}>
      {children}
    </div>
  );
  const NavyHdr = ({title,sub}:{title:string,sub?:string}) => (
    <div style={{ background:`linear-gradient(90deg,${C.navyMid},${C.navyLt})`, padding:"8px 16px", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
      <span style={{ color:C.goldLt, fontFamily:"'Sora',sans-serif", fontWeight:700, fontSize:13, textTransform:"uppercase" as const, letterSpacing:1 }}>{title}</span>
      {sub && <span style={{color:"#8899AA",fontSize:11}}>{sub}</span>}
    </div>
  );
  const RedHdr = ({title}:{title:string}) => (
    <div style={{ background:`linear-gradient(90deg,${C.mangRed},#9B1024)`, padding:"8px 16px" }}>
      <span style={{ color:"#fff", fontFamily:"'Sora',sans-serif", fontWeight:700, fontSize:13, textTransform:"uppercase" as const, letterSpacing:1 }}>{title}</span>
    </div>
  );
  const FRow = ({fkey}:{fkey:string}) => (
    <tr>
      <td style={{...tdL,width:280}}>{NCR_FILTER_LABELS[fkey]}</td>
      <td style={{...tdV, background:C.goldF, padding:"4px 10px"}}>
        <select value={filters[fkey]} onChange={e=>setF(fkey,e.target.value)} disabled={noData}
          style={{ width:"100%", border:"none", background:"transparent", fontFamily:"inherit", fontSize:13, fontWeight:700, color:C.blue, cursor:"pointer", outline:"none", textAlign:"center" as const }}>
          {(opts[fkey]||["ALL"]).map((o:string)=><option key={o} value={o}>{o}</option>)}
        </select>
      </td>
      <td style={tdN}>Dropdown</td>
    </tr>
  );
  const RRow = ({label,value,hl,big,last}:{label:string,value:any,hl?:boolean,big?:boolean,last?:boolean}) => {
    const isNum = typeof value==="number";
    const isNeg = isNum && value<0;
    const isPos = isNum && value>0;
    const valColor = isNeg ? C.red : isPos ? C.grn : C.ink;
    const bgColor  = hl ? (isNeg ? C.redF : C.grnF) : undefined;
    return (
    <tr>
      <td style={{...tdL, width:185, ...(last?{borderBottom:"none"}:{}), fontWeight:big?700:600}}>{label}</td>
      <td style={{ ...tdV, ...(last?{borderBottom:"none"}:{}), background:bgColor, color:valColor, fontSize:big?16:14, fontWeight:big?900:700 }}>
        {isNum?value.toLocaleString("en-IN"):value}
      </td>
      <td style={{...tdN,...(last?{borderBottom:"none"}:{})}}>← auto</td>
    </tr>
    );
  };

  return (
    <div>
      <style>{`
        .ncr-tbl{border-collapse:collapse;width:100%}
        .ncr-tbl td{vertical-align:middle}
        .ncr-tbl tbody tr:hover td{background:${C.goldF}!important;transition:background .12s}
        select{outline:none;-webkit-appearance:none;appearance:none}
        .ncr-ci{width:100px;border:1.5px solid ${C.border};border-radius:6px;padding:3px 8px;font-family:inherit;font-size:13px;text-align:right;outline:none;color:${C.ink};background:#fff}
        .ncr-ci:focus{border-color:${C.gold};box-shadow:0 0 0 2px rgba(212,137,10,.12)}
        .ncr-ubtn{background:linear-gradient(135deg,${C.goldLt},${C.gold});border:none;color:#fff;padding:8px 18px;border-radius:9px;font-family:inherit;font-size:13px;font-weight:700;cursor:pointer;transition:all .2s}
        .ncr-ubtn:hover{opacity:.88;transform:translateY(-1px)}
        .ncr-tbtn{background:rgba(245,184,0,.12);border:1px solid rgba(245,184,0,.4);color:${C.goldLt};padding:8px 16px;border-radius:9px;font-family:inherit;font-size:12px;font-weight:600;cursor:pointer}
        .ncr-tbtn:hover{background:rgba(245,184,0,.2)}
        .ncr-ebtn{background:rgba(245,184,0,.12);border:1px solid rgba(245,184,0,.35);color:${C.goldLt};padding:4px 12px;border-radius:6px;font-family:inherit;font-size:11px;font-weight:700;cursor:pointer}
        .ncr-ebtn:hover{background:rgba(245,184,0,.25)}
        .ncr-rbtn{background:none;border:1px solid ${C.grnBd};color:#065F46;padding:2px 10px;border-radius:6px;cursor:pointer;font-size:11px;font-weight:600;font-family:inherit}
        .ncr-pi{width:100%;border:none;background:transparent;font-weight:800;font-size:15px;color:${C.ink};text-align:right;outline:none;font-family:inherit;padding:6px 14px}
        .ncr-previ{width:100%;border:none;background:transparent;font-size:14px;color:${C.ink};text-align:right;outline:none;font-family:inherit;padding:4px 8px;font-weight:700}
      `}</style>

      {/* PAGE HEADER */}
      <div className="ph">
        <div>
          <div className="ph-eyebrow">Analytics</div>
          <div className="ph-title">NCR Calculator</div>
          <div className="ph-sub">Upload SAP billing data to analyse Net Contribution Rate across Trade / Non-Trade / NCM segments</div>
        </div>
        <div style={{display:"flex",gap:10,alignItems:"center"}}>
          {user?.role==="Admin" && <>
            <button className="ncr-tbtn" onClick={dlTemplate}>⬇ Download Template</button>
            <button className="ncr-ubtn" onClick={()=>fileRef.current?.click()}>📤 Upload Billing Data</button>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={handleFile}/>
          </>}
        </div>
      </div>

      {/* STATUS BAR */}
      {!noData && (
        <div style={{ background:C.grnF, border:`1.5px solid ${C.grnBd}`, borderRadius:10, padding:"8px 18px", display:"flex",alignItems:"center",gap:10, fontSize:12,color:"#065F46", marginBottom:16 }}>
          <span>✅</span>
          <span style={{fontWeight:700}}>{fileName}</span>
          <span style={{color:C.muted}}>·</span>
          <span>{rows.length} total rows</span>
          <span style={{color:C.muted}}>·</span>
          <span style={{fontWeight:700}}>{filtered.length} rows match current filters</span>
          <div style={{flex:1}}/>
          <span style={{color:C.muted}}>Data month:</span>
          <span style={{fontWeight:800,background:C.white,border:`1px solid ${C.grnBd}`,borderRadius:6,padding:"1px 10px"}}>{agg.lastMonth}</span>
          {user?.role==="Admin" && <button className="ncr-rbtn" onClick={()=>fileRef.current?.click()}>🔁 Replace File</button>}
        </div>
      )}

      {/* INSTRUCTION BANNER */}
      <div style={{ fontSize:11.5,color:C.muted,background:C.cream,borderRadius:8,border:`1px solid ${C.border2}`,padding:"6px 16px",marginBottom:14,fontStyle:"italic" }}>
        Enter inputs in blue cells · All values ₹ per MT · Costs auto-pull from COST SHEET history · Upload SAP billing Excel to activate filters
      </div>

      {noData ? (
        <div style={{ maxWidth:560, margin:"32px auto", background:C.white, border:`2px dashed ${C.border}`, borderRadius:18, padding:"42px 36px", textAlign:"center", boxShadow:"0 4px 24px rgba(180,130,0,.07)" }}>
          <div style={{fontSize:52,marginBottom:14}}>📊</div>
          <div style={{ fontFamily:"'Sora',sans-serif", fontSize:20, fontWeight:800, color:C.ink, marginBottom:8 }}>Upload Billing Data to Begin</div>
          <div style={{ color:C.muted, fontSize:13, marginBottom:16, lineHeight:1.7 }}>
            Upload your SAP billing Excel export to calculate NCR across Trade / Non-Trade / NCM with full multi-level filter support.
          </div>
          <div style={{ color:C.ink3, fontSize:11.5, background:C.cream, borderRadius:10, padding:"12px 16px", textAlign:"left", marginBottom:24, lineHeight:2.1, border:`1px solid ${C.border2}` }}>
            <strong style={{color:C.gold}}>Required Excel Columns:</strong><br/>
            MONTH · DistributionChannel · Ship Zone · Sold-topartyCode · Sold-topartyNAME · Grade · Ship City · SHP Sales District / Price Block / Taluk · MODE · UNIT · Source · Storagelocation · <strong>DIRECT DEPOT</strong> (Direct / Retailer) · <strong>TR &amp; NT</strong> (Trade / Non-Trade / NCM) · <strong>BillingQuantity</strong> · <strong>REALSN AMOUNT</strong> (₹/MT) · <strong>PRICE</strong> (₹/Bag)
          </div>
          <div style={{display:"flex",gap:12,justifyContent:"center"}}>
            {user?.role==="Admin" ? <>
              <button className="ncr-tbtn" onClick={dlTemplate} style={{padding:"10px 22px",borderRadius:10,fontSize:13,border:`1.5px solid ${C.gold}`,color:C.gold,background:C.goldF}}>⬇ Download Template</button>
              <button className="ncr-ubtn" onClick={()=>fileRef.current?.click()} style={{padding:"10px 22px",borderRadius:10,fontSize:13}}>📤 Upload Excel File</button>
            </> : <div style={{fontSize:13,color:C.muted,fontStyle:"italic",padding:"10px 0"}}>📊 Billing data is uploaded by Admin. Please contact your administrator.</div>}
          </div>
        </div>
      ) : (
        <>
          {/* MAIN GRID: LEFT + RIGHT */}
          <div style={{ display:"grid", gridTemplateColumns:"1fr 390px", gap:16, marginBottom:16, alignItems:"start" }}>

            {/* LEFT PANEL */}
            <div style={{display:"flex",flexDirection:"column",gap:14}}>
              {/* Previous Month */}
              <NCard>
                <NavyHdr title="Previous Month Calculation"/>
                <table className="ncr-tbl">
                  <tbody>
                    {Object.keys(NCR_FILTER_LABELS).map(fk=><FRow key={fk} fkey={fk}/>)}
                    <tr><td colSpan={3} style={{padding:0,height:3,background:`linear-gradient(90deg,${C.goldLt}55,transparent)`}}/></tr>
                    <tr>
                      <td style={{...tdL,width:280}}>Price (₹ / Bag)</td>
                      <td style={{...tdV,background:C.cream,fontWeight:700}}>{prevCalc.price||"—"}</td>
                      <td style={tdN}>← auto from data</td>
                    </tr>
                    <tr>
                      <td style={tdL}>GST @ 18% (₹ / Bag)</td>
                      <td style={{...tdV,background:C.cream,color:C.ink3}}>{prevCalc.gst||"—"}</td>
                      <td style={tdN}>← auto from data</td>
                    </tr>
                    <tr>
                      <td style={tdL}>Price Net of GST (₹ / Bag)</td>
                      <td style={{...tdV,background:C.cream,color:C.ink3}}>{prevCalc.net||"—"}</td>
                      <td style={tdN}>← auto from data</td>
                    </tr>
                    <tr>
                      <td style={{...tdL,fontWeight:800,fontSize:14,borderBottom:"none"}}>NCR / MT (₹ per Ton)</td>
                      <td style={{...tdV,borderBottom:"none",fontSize:19,fontWeight:900, background:prevCalc.ncr<0?C.redF:C.grnF, color:prevCalc.ncr<0?C.red:C.grn}}>
                        {prevCalc.ncr!=null?prevCalc.ncr.toLocaleString("en-IN"):"—"}
                      </td>
                      <td style={{...tdN,borderBottom:"none"}}>← auto from data</td>
                    </tr>
                  </tbody>
                </table>
              </NCard>

              {/* Current Month */}
              <NCard>
                <NavyHdr title="Current Month Calculation" sub="Per Bag (₹)"/>
                <table className="ncr-tbl">
                  <tbody>
                    <tr>
                      <td style={{...tdL,fontWeight:700,width:280}}>Current Price (₹ / Bag)</td>
                      <td style={{...tdV,background:"#FFFDE7",padding:0}}>
                        <input className="ncr-pi" type="number" value={currPx} onChange={e=>setCurrPx(e.target.value)} placeholder="Enter ₹ per bag…"/>
                      </td>
                      <td style={{...tdN,color:"#5A7080",fontStyle:"italic"}}>Manual</td>
                    </tr>
                    <tr>
                      <td style={tdL}>GST @ 18% (₹ / Bag)</td>
                      <td style={tdV}>{currPx?currCalc.gst:"—"}</td>
                      <td style={tdN}>← auto</td>
                    </tr>
                    <tr>
                      <td style={tdL}>Current Price Net of GST (₹ / Bag)</td>
                      <td style={tdV}>{currPx?currCalc.net:"—"}</td>
                      <td style={tdN}>← auto</td>
                    </tr>
                    <tr>
                      <td style={{...tdL,fontWeight:800,fontSize:14,borderBottom:"none"}}>NCR / MT (₹ per Ton)</td>
                      <td style={{...tdV,fontSize:19,fontWeight:900,borderBottom:"none",
                        background:!currPx?undefined:currCalc.ncr!=null&&currCalc.ncr<0?C.redF:C.grnF,
                        color:!currPx?C.muted:currCalc.ncr!=null&&currCalc.ncr<0?C.red:C.grn}}>
                        {currPx&&currCalc.ncr!=null?currCalc.ncr.toLocaleString("en-IN"):"—"}
                      </td>
                      <td style={{...tdN,borderBottom:"none"}}>← auto</td>
                    </tr>
                  </tbody>
                </table>
              </NCard>
            </div>

            {/* RIGHT PANEL */}
            <div style={{display:"flex",flexDirection:"column",gap:12}}>
              <NCard>
                <RedHdr title="Trade"/>
                <table className="ncr-tbl">
                  <tbody>
                    <tr>
                      <td style={{...tdL,width:185}}>Last Month (auto)</td>
                      <td style={{...tdV,background:C.grnF,color:"#065F46",fontWeight:800}}>{agg.lastMonth}</td>
                      <td style={tdN}>← auto</td>
                    </tr>
                    <RRow label="Billing Quantity"         value={agg.trade.qty}/>
                    <RRow label="TR Realization (₹ / MT)" value={agg.trade.realization} hl big/>
                    <RRow label="Direct % Ratio"           value={`${agg.trade.directPct}%`}/>
                    <RRow label="Retailer %"               value={`${agg.trade.retailerPct}%`}/>
                    <RRow label="Price (₹ / Bag)"          value={agg.trade.price} last/>
                  </tbody>
                </table>
              </NCard>
              <NCard>
                <RedHdr title="Non-Trade"/>
                <table className="ncr-tbl">
                  <tbody>
                    <RRow label="Billing Quantity"       value={agg.nt.qty}/>
                    <RRow label="Realization (₹ / MT)"  value={agg.nt.realization}  hl={agg.nt.qty>0} big={agg.nt.qty>0}/>
                    <RRow label="Price (₹ / Bag)"        value={agg.nt.price} last/>
                  </tbody>
                </table>
              </NCard>
              <NCard>
                <RedHdr title="NCM"/>
                <table className="ncr-tbl">
                  <tbody>
                    <RRow label="Billing Quantity"       value={agg.ncm.qty}/>
                    <RRow label="Realization (₹ / MT)"  value={agg.ncm.realization} hl={agg.ncm.qty>0} big={agg.ncm.qty>0}/>
                    <RRow label="Price (₹ / Bag)"        value={agg.ncm.price} last/>
                  </tbody>
                </table>
              </NCard>
              <NCard>
                <RedHdr title="Overall"/>
                <table className="ncr-tbl">
                  <tbody>
                    <RRow label="Billing Quantity"       value={agg.overall.qty}/>
                    <RRow label="Realization (₹ / MT)"  value={agg.overall.realization} hl big/>
                    <RRow label="Price (₹ / Bag)"        value={agg.overall.price} last/>
                  </tbody>
                </table>
              </NCard>
              <div style={{ background:C.grnF, border:`1.5px solid ${C.grnBd}`, borderRadius:10, padding:"10px 16px", display:"flex",justifyContent:"space-between",alignItems:"center" }}>
                <span style={{fontSize:12,fontWeight:600,color:"#065F46"}}>Matching rows (full filter)</span>
                <span style={{fontSize:22,fontWeight:900,color:C.grn,fontFamily:"'Sora',sans-serif"}}>{filtered.length}</span>
              </div>
            </div>
          </div>

          {/* COST COMPONENTS TABLE */}
          <NCard>
            <div style={{ background:`linear-gradient(90deg,${C.navyMid},${C.navyLt})`, padding:"9px 16px", display:"flex",alignItems:"center",justifyContent:"space-between" }}>
              <span style={{ color:C.goldLt, fontFamily:"'Sora',sans-serif", fontWeight:700, fontSize:13, textTransform:"uppercase" as const, letterSpacing:1 }}>Cost Components</span>
              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                <span style={{color:"#8899AA",fontSize:11}}>Costs auto-pull from COST SHEET history · Edit to override</span>
                <button className="ncr-ebtn" onClick={()=>setEditC(x=>!x)}>{editC?"✅ Done":"✏️ Edit"}</button>
              </div>
            </div>
            <div style={{overflowX:"auto"}}>
              <table className="ncr-tbl" style={{tableLayout:"fixed" as const}}>
                <thead>
                  <tr>
                    {[["40%","Cost Components"],["30%","Previous Month Per MT (₹)"],["30%","Current Month Per MT (₹)"]].map(([w,l])=>(
                      <th key={l} style={{ background:C.navyMid, color:C.goldLt, padding:"9px 16px", fontSize:11, fontWeight:700, textTransform:"uppercase" as const, letterSpacing:.7, width:w, textAlign:(l==="Cost Components"?"left":"right") as any }}>{l}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {NCR_COST_KEYS.map(([key,label])=>(
                    <tr key={key}>
                      <td style={{...tdL,color:C.gold,fontWeight:700}}>{label}</td>
                      <td style={{...tdV,color:C.grn}}>
                        {costs[key].prev.toFixed(2)}
                      </td>
                      <td style={{...tdV,color:C.grn}}>
                        {editC?<input className="ncr-ci" type="number" value={costs[key].curr} onChange={e=>setCost(key,"curr",e.target.value)}/>:costs[key].curr.toFixed(2)}
                      </td>
                    </tr>
                  ))}
                  <tr style={{background:C.goldF}}>
                    <td style={{...tdL,fontWeight:800,color:C.ink,fontSize:13.5}}>Total Expenses</td>
                    <td style={{...tdV,fontWeight:800,color:C.grn,fontSize:13.5}}>{prevTotalExp.toFixed(2)}</td>
                    <td style={{...tdV,fontWeight:800,color:C.grn,fontSize:13.5}}>{currTotalExp.toFixed(2)}</td>
                  </tr>
                  <tr>
                    <td style={{...tdL,color:C.ink3,borderBottom:"none"}}>Total Cost of Production</td>
                    <td style={{...tdV,color:C.ink3,borderBottom:"none"}}>
                      {costs.cop.prev.toFixed(2)}
                    </td>
                    <td style={{...tdV,color:C.ink3,borderBottom:"none"}}>
                      {editC?<input className="ncr-ci" type="number" value={costs.cop.curr} onChange={e=>setCosts(c=>({...c,cop:{...c.cop,curr:parseFloat(e.target.value)||0}}))}/>:costs.cop.curr.toFixed(2)}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </NCard>

          {/* Formula reference */}
          <div style={{ marginTop:12, background:C.cream, border:`1px solid ${C.border2}`, borderRadius:10, padding:"10px 18px", display:"flex", gap:24, fontSize:11.5, color:C.muted, flexWrap:"wrap" as const }}>
            <span>📐 <strong>Previous NCR/MT</strong> = REALSN AMOUNT ÷ BillingQuantity (weighted avg from COST SHEET)</span>
            <span>📐 <strong>Current NCR/MT</strong> = (CurrentPrice ÷ 1.18) × 20 − Total Expenses − CoP</span>
            <span>📐 <strong>Total Expenses</strong> = Sum of all 21 cost components above (auto-pulled from COST SHEET)</span>
            <span>📐 <strong>1 MT = 20 bags @ 50 kg</strong></span>
          </div>
        </>
      )}
    </div>
  );
}


// ═══ PRICE MASTER ════════════════════════════════════════════
const INITIAL_PRICE_MASTER = [
  {id:"PM001",condType:"ZLCO",regionCode:4,regionName:"Bihar",cityCode:402,cityName:"MUZAFFARPUR",salesDistrict:40201,salesDistrictName:"MUZAFFARPUR",validFrom:"2025-09-22",grade:"PPC",rate:323,validTo:"9999-12-31"},
  {id:"PM002",condType:"ZLCO",regionCode:4,regionName:"Bihar",cityCode:402,cityName:"MUZAFFARPUR",salesDistrict:40202,salesDistrictName:"MUSHARI",validFrom:"2025-09-22",grade:"PPC",rate:327,validTo:"9999-12-31"},
  {id:"PM003",condType:"ZLCO",regionCode:24,regionName:"Uttar Pradesh",cityCode:2427,cityName:"SHAMLI",salesDistrict:242704,salesDistrictName:"BUDHANA",validFrom:"2025-09-22",grade:"PPC",rate:348,validTo:"9999-12-31"},
  {id:"PM004",condType:"ZLCO",regionCode:24,regionName:"Uttar Pradesh",cityCode:2428,cityName:"KANPUR",salesDistrict:242801,salesDistrictName:"KANPUR",validFrom:"2025-09-22",grade:"PPC",rate:308.5,validTo:"9999-12-31"},
  {id:"PM005",condType:"ZLCO",regionCode:24,regionName:"Uttar Pradesh",cityCode:2430,cityName:"JHANSI",salesDistrict:243001,salesDistrictName:"JHANSI",validFrom:"2025-09-22",grade:"PPC",rate:320,validTo:"9999-12-31"},
  {id:"PM006",condType:"ZLCO",regionCode:20,regionName:"Rajasthan",cityCode:2001,cityName:"JAIPUR",salesDistrict:200101,salesDistrictName:"JAIPUR",validFrom:"2025-09-22",grade:"PPC",rate:340,validTo:"9999-12-31"},
  {id:"PM007",condType:"ZLCO",regionCode:20,regionName:"Rajasthan",cityCode:2002,cityName:"KOTA",salesDistrict:200201,salesDistrictName:"KOTA",validFrom:"2025-09-22",grade:"PPC",rate:335,validTo:"9999-12-31"},
  {id:"PM008",condType:"ZLCO",regionCode:20,regionName:"Rajasthan",cityCode:2001,cityName:"JAIPUR",salesDistrict:200101,salesDistrictName:"JAIPUR",validFrom:"2025-09-22",grade:"OPC",rate:360,validTo:"9999-12-31"},
];

const INITIAL_FREIGHT_MASTER = [
  {id:"FM001",sendingZone:"1201010002",from:"BASAI (MANDSAUR)",sendState:"Madhya Pradesh",receivingZone:"2020060006",to:"DABLI RATHAN",recvState:"Rajasthan",district:202006,districtName:"HANUMANGARH",route:"ST1031",routeDesc:"HANUMANGARH TO DABLI RATHAN",freightPMT:144,distance:17,freightPerKm:8.47,validFrom:"2016-04-01",validTo:"9999-12-31"},
  {id:"FM002",sendingZone:"1201010002",from:"BASAI (MANDSAUR)",sendState:"Madhya Pradesh",receivingZone:"2020020002",to:"DEENGARH",recvState:"Rajasthan",district:202002,districtName:"SANGRIA",route:"ST1032",routeDesc:"HANUMANGARH TO DEENGARH",freightPMT:0,distance:38,freightPerKm:0,validFrom:"2023-01-09",validTo:"9999-12-31"},
  {id:"FM003",sendingZone:"1201010002",from:"BASAI (MANDSAUR)",sendState:"Madhya Pradesh",receivingZone:"2020020003",to:"DHABAN",recvState:"Rajasthan",district:202002,districtName:"SANGRIA",route:"ST1033",routeDesc:"HANUMANGARH TO DHABAN",freightPMT:215,distance:45,freightPerKm:4.78,validFrom:"2016-04-01",validTo:"9999-12-31"},
  {id:"FM004",sendingZone:"1201010002",from:"BASAI (MANDSAUR)",sendState:"Madhya Pradesh",receivingZone:"2020060009",to:"DHALIYA",recvState:"Rajasthan",district:202006,districtName:"HANUMANGARH",route:"ST1034",routeDesc:"HANUMANGARH TO DHALIYA",freightPMT:130,distance:15,freightPerKm:8.67,validFrom:"2016-04-01",validTo:"9999-12-31"},
  {id:"FM005",sendingZone:"1201010002",from:"BASAI (MANDSAUR)",sendState:"Madhya Pradesh",receivingZone:"2020060013",to:"FATEHGARH (HMO)",recvState:"Rajasthan",district:202006,districtName:"HANUMANGARH",route:"ST1036",routeDesc:"HANUMANGARH TO FATEHGARH",freightPMT:0,distance:30,freightPerKm:0,validFrom:"2023-01-09",validTo:"9999-12-31"},
];

const INITIAL_LOCATION_MASTER = {
  countries:  [{id:"C1",name:"India"}],
  states:     [{id:"S1",name:"Rajasthan",countryId:"C1"},{id:"S2",name:"Uttar Pradesh",countryId:"C1"},{id:"S3",name:"Madhya Pradesh",countryId:"C1"},{id:"S4",name:"Bihar",countryId:"C1"}],
  regions:    [{id:"R1",name:"RAJ (Rajasthan)",stateId:"S1"},{id:"R2",name:"NZ (North Zone)",stateId:"S2"},{id:"R3",name:"MP (Madhya Pradesh)",stateId:"S3"}],
  clusters:   [{id:"CL1",name:"Raj-1 (Kota)",regionId:"R1"},{id:"CL2",name:"Raj-2 (Jaipur)",regionId:"R1"},{id:"CL3",name:"NZ-1",regionId:"R2"}],
  salesOffices:[{id:"SO1",name:"KOTA SO",clusterId:"CL1"},{id:"SO2",name:"JAIPUR SO",clusterId:"CL2"},{id:"SO3",name:"AGRA SO",clusterId:"CL3"}],
  districts:  [{id:"DT1",name:"KOTA",salesOfficeId:"SO1"},{id:"DT2",name:"JAIPUR",salesOfficeId:"SO2"},{id:"DT3",name:"AJMER",salesOfficeId:"SO2"}],
  tehsils:    [{id:"TH1",name:"LADPURA",districtId:"DT1"},{id:"TH2",name:"PIPALDA",districtId:"DT1"},{id:"TH3",name:"SANGANER",districtId:"DT2"}],
  cities:     [{id:"CT1",name:"KOTA CITY",tehsilId:"TH1"},{id:"CT2",name:"NAYAPURA",tehsilId:"TH1"},{id:"CT3",name:"JAIPUR CITY",tehsilId:"TH3"}],
};

// ─── Price Master Panel ───────────────────────────────────────
function PriceMasterPanel({ priceMaster, setPriceMaster }) {
  const [search, setSearch] = useState("");
  const [fState, setFState] = useState(""); const [fGrade, setFGrade] = useState("");
  const [showAdd, setShowAdd] = useState(false);
  const [editId, setEditId] = useState(null);
  const [form, setForm] = useState({condType:"ZLCO",regionCode:"",regionName:"",cityCode:"",cityName:"",salesDistrict:"",salesDistrictName:"",validFrom:"",grade:"PPC",rate:"",validTo:"9999-12-31"});
  const fileRef = useRef();

  const handleExcel = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, {type:"array"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, {defval:""});
        const mapped = rows.filter(r=>r["Rate"]!==undefined).map((r,i)=>({
          id:`PM${Date.now()}_${i}`,
          condType: String(r["Condition type"]||"ZLCO"),
          regionCode: Number(r["Region"]||0),
          regionName: String(r["Region Name"]||""),
          cityCode: Number(r["City Code"]||0),
          cityName: String(r["City Name"]||""),
          salesDistrict: Number(r["Sales District"]||0),
          salesDistrictName: String(r["Sales District Name"]||""),
          validFrom: r["Validity start date"]?String(r["Validity start date"]).split("T")[0]:"",
          materialPriceGroup: String(r["Material Price Group"]||r["Material Price Grp"]||""),
          grade: String(r["Material Description"]||"PPC"),
          rate: Number(r["Rate"]||0),
          validTo: String(r["Validity End Date"]||"9999-12-31").split("T")[0],
        }));
        setPriceMaster(mapped);
        alert(`✅ Loaded ${mapped.length} price records`);
      } catch(err) { alert("Error: "+err.message); }
    };
    reader.readAsArrayBuffer(file); e.target.value="";
  };

  const states = [...new Set(priceMaster.map(r=>r.regionName).filter(Boolean))].sort();
  const grades  = [...new Set(priceMaster.map(r=>r.grade).filter(Boolean))].sort();
  const filtered = priceMaster.filter(r => {
    if (fState && r.regionName !== fState) return false;
    if (fGrade && r.grade !== fGrade) return false;
    if (search) { const s = search.toLowerCase(); return [r.regionName,r.cityName,r.salesDistrictName,r.grade].some(v=>String(v).toLowerCase().includes(s)); }
    return true;
  });

  const save = () => {
    if (editId) { setPriceMaster(p=>p.map(r=>r.id===editId?{...r,...form,id:editId}:r)); setEditId(null); }
    else { setPriceMaster(p=>[...p,{...form,id:`PM${Date.now()}`,regionCode:Number(form.regionCode),cityCode:Number(form.cityCode),salesDistrict:Number(form.salesDistrict),rate:Number(form.rate)}]); }
    setShowAdd(false); setForm({condType:"ZLCO",regionCode:"",regionName:"",cityCode:"",cityName:"",salesDistrict:"",salesDistrictName:"",validFrom:"",grade:"PPC",rate:"",validTo:"9999-12-31"});
  };
  const edit = (r) => { setForm({...r}); setEditId(r.id); setShowAdd(true); };
  const del  = (id) => { if(confirm("Delete this price record?")) setPriceMaster(p=>p.filter(r=>r.id!==id)); };

  return (
    <div>
      <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:16,flexWrap:"wrap"}}>
        <input className="fc" style={{flex:1,minWidth:200}} placeholder="🔍 Search state, city, district, grade…" value={search} onChange={e=>setSearch(e.target.value)}/>
        <select className="fc" style={{width:160}} value={fState} onChange={e=>setFState(e.target.value)}><option value="">All States</option>{states.map(s=><option key={s}>{s}</option>)}</select>
        <select className="fc" style={{width:120}} value={fGrade} onChange={e=>setFGrade(e.target.value)}><option value="">All Grades</option>{grades.map(g=><option key={g}>{g}</option>)}</select>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={handleExcel}/>
        <button className="btn btn-secondary btn-sm" onClick={()=>fileRef.current?.click()}>📤 Upload Excel</button>
        <button className="btn btn-primary btn-sm" onClick={()=>{setShowAdd(true);setEditId(null);}}>+ Add Row</button>
      </div>
      <div style={{fontSize:12,color:"var(--muted)",marginBottom:10}}>{filtered.length} of {priceMaster.length} records</div>

      {showAdd && (
        <div className="card" style={{marginBottom:16,border:"2px solid var(--gold-200)"}}>
          <div className="card-hdr"><div className="card-title">{editId?"Edit":"Add"} Price Record</div><button className="btn btn-ghost btn-sm" onClick={()=>{setShowAdd(false);setEditId(null);}}>✕</button></div>
          <div className="card-body">
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:12}}>
              {[["regionCode","Region Code"],["regionName","Region / State"],["cityCode","City Code"],["cityName","City Name"],["salesDistrict","Sales District"],["salesDistrictName","Sales District Name"]].map(([k,l])=>(
                <div className="fg" key={k}><label className="fl">{l}</label><input className="fc" value={form[k]} onChange={e=>setForm(p=>({...p,[k]:e.target.value}))}/></div>
              ))}
              <div className="fg"><label className="fl">Grade</label><select className="fc" value={form.grade} onChange={e=>setForm(p=>({...p,grade:e.target.value}))}><option>PPC</option><option>OPC</option><option>PMX</option></select></div>
              <div className="fg"><label className="fl">Rate (₹/Bag)</label><input className="fc" type="number" value={form.rate} onChange={e=>setForm(p=>({...p,rate:e.target.value}))}/></div>
              <div className="fg"><label className="fl">Valid From</label><input className="fc" type="date" value={form.validFrom} onChange={e=>setForm(p=>({...p,validFrom:e.target.value}))}/></div>
              <div className="fg"><label className="fl">Valid To</label><input className="fc" value={form.validTo} onChange={e=>setForm(p=>({...p,validTo:e.target.value}))}/></div>
            </div>
            <div style={{display:"flex",gap:8,marginTop:12}}>
              <button className="btn btn-primary btn-sm" onClick={save}>💾 Save</button>
              <button className="btn btn-ghost btn-sm" onClick={()=>{setShowAdd(false);setEditId(null);}}>Cancel</button>
            </div>
          </div>
        </div>
      )}

      <div className="tw"><table>
        <thead><tr><th>Condition Type</th><th>Region</th><th>Region Name</th><th>City Code</th><th>City Name</th><th>Sales District</th><th>Sales District Name</th><th>Matl.Price Grp</th><th>Material Description</th><th style={{textAlign:"right"}}>Rate (₹)</th><th>Valid From</th><th>Valid To</th><th></th></tr></thead>
        <tbody>
          {filtered.length===0 && <tr><td colSpan={13} style={{textAlign:"center",padding:40,color:"var(--muted)"}}>No records. Upload Excel or add rows.</td></tr>}
          {filtered.map(r=>(
            <tr key={r.id}>
              <td style={{fontSize:11,color:"var(--muted)"}}>{r.condType}</td>
              <td style={{fontWeight:700,fontSize:12}}>{r.regionCode}</td>
              <td>{r.regionName}</td>
              <td style={{color:"var(--muted)",fontSize:12}}>{r.cityCode}</td>
              <td style={{fontWeight:600}}>{r.cityName}</td>
              <td style={{color:"var(--muted)",fontSize:12}}>{r.salesDistrict}</td>
              <td>{r.salesDistrictName}</td>
              <td style={{textAlign:"center",color:"var(--muted)",fontSize:12}}>{r.materialPriceGroup||"—"}</td>
              <td><span className="badge b-grade">{r.grade}</span></td>
              <td style={{textAlign:"right",fontFamily:"'Sora',sans-serif",fontWeight:800,color:r.rate>0?"var(--grn)":"var(--muted)"}}>{r.rate>0?`₹${r.rate}`:"—"}</td>
              <td style={{fontSize:11.5,color:"var(--muted)"}}>{r.validFrom}</td>
              <td style={{fontSize:11.5,color:"var(--muted)"}}>{r.validTo}</td>
              <td><div style={{display:"flex",gap:4}}><button className="btn btn-ghost btn-sm" style={{fontSize:11}} onClick={()=>edit(r)}>✏️</button><button className="btn btn-ghost btn-sm" style={{fontSize:11,color:"var(--red)"}} onClick={()=>del(r.id)}>🗑</button></div></td>
            </tr>
          ))}
        </tbody>
      </table></div>
    </div>
  );
}

// ─── Freight Master Panel ─────────────────────────────────────
function FreightMasterPanel({ freightMaster, setFreightMaster }) {
  const [search, setSearch] = useState("");
  const [fFrom, setFFrom] = useState(""); const [fState, setFState] = useState("");
  const [showAdd, setShowAdd] = useState(false); const [editId, setEditId] = useState(null);
  const [form, setForm] = useState({from:"",sendState:"",to:"",recvState:"",districtName:"",route:"",routeDesc:"",freightPMT:"",distance:"",freightPerKm:"",validFrom:"",validTo:"9999-12-31"});
  const fileRef = useRef();

  const handleExcel = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, {type:"array"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, {defval:""});
        const mapped = rows.filter(r=>r["Route"]).map((r,i)=>({
          id:`FM${Date.now()}_${i}`,
          sendingZone: String(r["Sending Zone"]||""),
          from:        String(r["From"]||""),
          sendState:   String(r["Send.State Name"]||""),
          receivingZone: String(r["Receving Zone"]||""),
          to:          String(r["To"]||""),
          recvState:   String(r["Recv.State Name"]||""),
          district:    Number(r["District"]||0),
          districtName:String(r["District Name"]||""),
          route:       String(r["Route"]||""),
          routeDesc:   String(r["Description of Route"]||""),
          freightPMT:  Number(r["Freight PMT"]||0),
          distance:    Number(r["Distance"]||0),
          freightPerKm:Number(r["Freight PMT/KM"]||0),
          validFrom:   r["Validity Start Date"]?String(r["Validity Start Date"]).split("T")[0]:"",
          validTo:     r["Validity End Date"]?String(r["Validity End Date"]).split("T")[0]:"9999-12-31",
        }));
        setFreightMaster(mapped);
        alert(`✅ Loaded ${mapped.length} freight records`);
      } catch(err) { alert("Error: "+err.message); }
    };
    reader.readAsArrayBuffer(file); e.target.value="";
  };

  const fromList  = [...new Set(freightMaster.map(r=>r.from).filter(Boolean))].sort();
  const stateList = [...new Set(freightMaster.map(r=>r.recvState).filter(Boolean))].sort();
  const filtered  = freightMaster.filter(r=>{
    if (fFrom  && r.from!==fFrom)           return false;
    if (fState && r.recvState!==fState)     return false;
    if (search){ const s=search.toLowerCase(); return [r.from,r.to,r.route,r.routeDesc,r.districtName].some(v=>String(v).toLowerCase().includes(s)); }
    return true;
  });

  const save = () => {
    if (editId) { setFreightMaster(p=>p.map(r=>r.id===editId?{...r,...form,id:editId}:r)); setEditId(null); }
    else { setFreightMaster(p=>[...p,{...form,id:`FM${Date.now()}`,freightPMT:Number(form.freightPMT),distance:Number(form.distance),freightPerKm:Number(form.freightPerKm)}]); }
    setShowAdd(false); setForm({from:"",sendState:"",to:"",recvState:"",districtName:"",route:"",routeDesc:"",freightPMT:"",distance:"",freightPerKm:"",validFrom:"",validTo:"9999-12-31"});
  };
  const edit = (r) => { setForm({...r}); setEditId(r.id); setShowAdd(true); };
  const del  = (id) => { if(confirm("Delete this freight record?")) setFreightMaster(p=>p.filter(r=>r.id!==id)); };

  return (
    <div>
      <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:16,flexWrap:"wrap"}}>
        <input className="fc" style={{flex:1,minWidth:200}} placeholder="🔍 Search route, from, to, district…" value={search} onChange={e=>setSearch(e.target.value)}/>
        <select className="fc" style={{width:200}} value={fFrom} onChange={e=>setFFrom(e.target.value)}><option value="">All Source Plants</option>{fromList.map(s=><option key={s}>{s}</option>)}</select>
        <select className="fc" style={{width:160}} value={fState} onChange={e=>setFState(e.target.value)}><option value="">All Recv. States</option>{stateList.map(s=><option key={s}>{s}</option>)}</select>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={handleExcel}/>
        <button className="btn btn-secondary btn-sm" onClick={()=>fileRef.current?.click()}>📤 Upload Excel</button>
        <button className="btn btn-primary btn-sm" onClick={()=>{setShowAdd(true);setEditId(null);}}>+ Add Row</button>
      </div>
      <div style={{fontSize:12,color:"var(--muted)",marginBottom:10}}>{filtered.length} of {freightMaster.length} records</div>

      {showAdd && (
        <div className="card" style={{marginBottom:16,border:"2px solid var(--gold-200)"}}>
          <div className="card-hdr"><div className="card-title">{editId?"Edit":"Add"} Freight Record</div><button className="btn btn-ghost btn-sm" onClick={()=>{setShowAdd(false);setEditId(null);}}>✕</button></div>
          <div className="card-body">
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:12}}>
              {[["from","Source (From)"],["sendState","Sending State"],["to","Destination (To)"],["recvState","Receiving State"],["districtName","District Name"],["route","Route Code"],["routeDesc","Route Description"],["freightPMT","Freight PMT (₹)"],["distance","Distance (km)"],["freightPerKm","Freight/KM"],["validFrom","Valid From"],["validTo","Valid To"]].map(([k,l])=>(
                <div className="fg" key={k}><label className="fl">{l}</label><input className="fc" type={["freightPMT","distance","freightPerKm"].includes(k)?"number":"text"} value={form[k]||""} onChange={e=>setForm(p=>({...p,[k]:e.target.value}))}/></div>
              ))}
            </div>
            <div style={{display:"flex",gap:8,marginTop:12}}>
              <button className="btn btn-primary btn-sm" onClick={save}>💾 Save</button>
              <button className="btn btn-ghost btn-sm" onClick={()=>{setShowAdd(false);setEditId(null);}}>Cancel</button>
            </div>
          </div>
        </div>
      )}

      <div className="tw"><table>
        <thead><tr><th>Route</th><th>From (Source)</th><th>Send State</th><th>To (Destination)</th><th>Recv State</th><th>District</th><th>Route Desc.</th><th style={{textAlign:"right"}}>Freight PMT (₹)</th><th style={{textAlign:"right"}}>Dist (km)</th><th style={{textAlign:"right"}}>₹/KM</th><th>Valid From</th><th></th></tr></thead>
        <tbody>
          {filtered.length===0 && <tr><td colSpan={12} style={{textAlign:"center",padding:40,color:"var(--muted)"}}>No records. Upload Excel or add rows.</td></tr>}
          {filtered.map(r=>(
            <tr key={r.id}>
              <td style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:12,color:"var(--gold-dk)"}}>{r.route}</td>
              <td style={{fontWeight:600}}>{r.from}</td>
              <td style={{fontSize:12}}>{r.sendState}</td>
              <td style={{fontWeight:600}}>{r.to}</td>
              <td style={{fontSize:12}}>{r.recvState}</td>
              <td>{r.districtName}</td>
              <td style={{fontSize:11.5,color:"var(--muted)",maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.routeDesc}</td>
              <td style={{textAlign:"right",fontFamily:"'Sora',sans-serif",fontWeight:800,color:r.freightPMT>0?"var(--ink)":"var(--muted)"}}>{r.freightPMT>0?`₹${r.freightPMT}`:"—"}</td>
              <td style={{textAlign:"right",color:"var(--muted)"}}>{r.distance}</td>
              <td style={{textAlign:"right",color:"var(--muted)"}}>{r.freightPerKm}</td>
              <td style={{fontSize:11.5,color:"var(--muted)"}}>{r.validFrom}</td>
              <td><div style={{display:"flex",gap:4}}><button className="btn btn-ghost btn-sm" style={{fontSize:11}} onClick={()=>edit(r)}>✏️</button><button className="btn btn-ghost btn-sm" style={{fontSize:11,color:"var(--red)"}} onClick={()=>del(r.id)}>🗑</button></div></td>
            </tr>
          ))}
        </tbody>
      </table></div>
    </div>
  );
}

// ─── Location Master Panel ────────────────────────────────────
const LOC_LEVELS = ["countries","states","regions","clusters","salesOffices","districts","tehsils","cities"];
const LOC_LABELS = {countries:"Country",states:"State",regions:"Region",clusters:"Cluster",salesOffices:"Sales Office",districts:"District",tehsils:"Tehsil",cities:"City"};
const LOC_PARENT = {states:"countryId",regions:"stateId",clusters:"regionId",salesOffices:"clusterId",districts:"salesOfficeId",tehsils:"districtId",cities:"tehsilId"};

function LocationMasterPanel({ locationMaster, setLocationMaster, users, setUsers }) {
  const [lvl, setLvl] = useState("countries");
  const [search, setSearch] = useState("");
  const [showAdd, setShowAdd] = useState(false);
  const [editId, setEditId] = useState(null);
  const [form, setForm]   = useState({name:"",parentId:""});
  const [showRights, setShowRights] = useState(false);
  const [selUser, setSelUser]       = useState("");
  const [selSOs, setSelSOs]         = useState([]);

  const parentKey = LOC_PARENT[lvl];
  const parentLvl = LOC_LEVELS[LOC_LEVELS.indexOf(lvl)-1];

  const parentName = (parentId) => {
    if (!parentLvl) return "";
    const p = (locationMaster[parentLvl]||[]).find(x=>x.id===parentId);
    return p?p.name:"";
  };

  const items = (locationMaster[lvl]||[]).filter(r=>{
    if (search) return r.name.toLowerCase().includes(search.toLowerCase());
    return true;
  });

  const save = () => {
    const newItem = {id:`${lvl.slice(0,2).toUpperCase()}${Date.now()}`,name:form.name,...(parentKey?{[parentKey]:form.parentId}:{})};
    if (editId) {
      setLocationMaster(p=>({...p,[lvl]:p[lvl].map(r=>r.id===editId?{...newItem,id:editId}:r)}));
      setEditId(null);
    } else {
      setLocationMaster(p=>({...p,[lvl]:[...p[lvl],newItem]}));
    }
    setShowAdd(false); setForm({name:"",parentId:""});
  };
  const del = (id) => { if(confirm("Delete?")) setLocationMaster(p=>({...p,[lvl]:p[lvl].filter(r=>r.id!==id)})); };
  const edit = (r) => { setForm({name:r.name,parentId:r[parentKey]||""}); setEditId(r.id); setShowAdd(true); };

  // User location rights
  const salesOffices = locationMaster.salesOffices||[];
  const openRights = (u) => { setSelUser(u.id); setSelSOs(u.allowedSalesOffices||[]); setShowRights(true); };
  const saveRights = () => {
    setUsers(p=>p.map(u=>u.id===selUser?{...u,allowedSalesOffices:selSOs}:u));
    setShowRights(false);
  };
  const toggleSO = (soId) => setSelSOs(p=>p.includes(soId)?p.filter(x=>x!==soId):[...p,soId]);

  return (
    <div>
      {/* Level tabs */}
      <div className="tabs" style={{marginBottom:16}}>
        {LOC_LEVELS.map(l=><div key={l} className={`tab ${lvl===l?"active":""}`} onClick={()=>{setLvl(l);setSearch("");setShowAdd(false);setEditId(null);}}>{LOC_LABELS[l]}</div>)}
      </div>

      {/* User Location Rights sub-section */}
      <div className="card" style={{marginBottom:16,border:"1.5px solid var(--gold-200)"}}>
        <div className="card-hdr" style={{cursor:"pointer"}} onClick={()=>setShowRights(x=>!x)}>
          <div className="card-title">🔐 User Location Access Rights</div>
          <div style={{fontSize:12,color:"var(--muted)"}}>Assign Sales Office access per user</div>
        </div>
        {showRights && (
          <div className="card-body">
            <div className="fg" style={{marginBottom:12}}>
              <label className="fl">Select User</label>
              <select className="fc" style={{maxWidth:280}} value={selUser} onChange={e=>{setSelUser(e.target.value);const u=users.find(u=>u.id===e.target.value);setSelSOs(u?.allowedSalesOffices||[]);}}>
                <option value="">-- pick user --</option>
                {users.map(u=><option key={u.id} value={u.id}>{u.name} ({u.role})</option>)}
              </select>
            </div>
            {selUser && (
              <>
                <div style={{fontSize:13,fontWeight:700,marginBottom:8}}>Allowed Sales Offices:</div>
                <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:12}}>
                  {salesOffices.map(so=>(
                    <label key={so.id} style={{display:"flex",alignItems:"center",gap:6,padding:"6px 12px",border:`1.5px solid ${selSOs.includes(so.id)?"var(--gold)":"var(--border)"}`,borderRadius:8,cursor:"pointer",background:selSOs.includes(so.id)?"var(--gold-f)":"var(--white)",fontWeight:selSOs.includes(so.id)?700:400,fontSize:13}}>
                      <input type="checkbox" checked={selSOs.includes(so.id)} onChange={()=>toggleSO(so.id)} style={{accentColor:"var(--gold)"}}/>
                      {so.name}
                    </label>
                  ))}
                  {salesOffices.length===0 && <span style={{color:"var(--muted)",fontSize:13}}>Add Sales Offices in the Sales Office tab first.</span>}
                </div>
                {selSOs.length===0 && <div style={{fontSize:12,color:"var(--muted)",marginBottom:8}}>⚠️ No offices selected = user sees ALL data</div>}
                <button className="btn btn-primary btn-sm" onClick={saveRights}>💾 Save Rights</button>
              </>
            )}
          </div>
        )}
      </div>

      {/* Master CRUD */}
      <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:12}}>
        <input className="fc" style={{flex:1,maxWidth:300}} placeholder={`🔍 Search ${LOC_LABELS[lvl]}…`} value={search} onChange={e=>setSearch(e.target.value)}/>
        <button className="btn btn-primary btn-sm" onClick={()=>{setShowAdd(true);setEditId(null);setForm({name:"",parentId:""});}}>+ Add {LOC_LABELS[lvl]}</button>
      </div>

      {showAdd && (
        <div className="card" style={{marginBottom:12,border:"2px solid var(--gold-200)"}}>
          <div className="card-hdr"><div className="card-title">{editId?"Edit":"Add"} {LOC_LABELS[lvl]}</div><button className="btn btn-ghost btn-sm" onClick={()=>{setShowAdd(false);setEditId(null);}}>✕</button></div>
          <div className="card-body">
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
              <div className="fg"><label className="fl">{LOC_LABELS[lvl]} Name</label><input className="fc" value={form.name} onChange={e=>setForm(p=>({...p,name:e.target.value}))}/></div>
              {parentKey && parentLvl && (
                <div className="fg">
                  <label className="fl">{LOC_LABELS[parentLvl]}</label>
                  <select className="fc" value={form.parentId} onChange={e=>setForm(p=>({...p,parentId:e.target.value}))}>
                    <option value="">-- Select {LOC_LABELS[parentLvl]} --</option>
                    {(locationMaster[parentLvl]||[]).map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
                  </select>
                </div>
              )}
            </div>
            <div style={{display:"flex",gap:8,marginTop:12}}>
              <button className="btn btn-primary btn-sm" onClick={save}>💾 Save</button>
              <button className="btn btn-ghost btn-sm" onClick={()=>{setShowAdd(false);setEditId(null);}}>Cancel</button>
            </div>
          </div>
        </div>
      )}

      <div className="tw"><table>
        <thead><tr><th>ID</th><th>{LOC_LABELS[lvl]} Name</th>{parentKey&&<th>Parent ({LOC_LABELS[parentLvl]})</th>}<th></th></tr></thead>
        <tbody>
          {items.length===0&&<tr><td colSpan={4} style={{textAlign:"center",padding:32,color:"var(--muted)"}}>No {LOC_LABELS[lvl]} entries yet.</td></tr>}
          {items.map(r=>(
            <tr key={r.id}>
              <td style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:12,color:"var(--gold-dk)"}}>{r.id}</td>
              <td style={{fontWeight:700}}>{r.name}</td>
              {parentKey&&<td style={{color:"var(--muted)",fontSize:12}}>{parentName(r[parentKey])}</td>}
              <td><div style={{display:"flex",gap:4}}><button className="btn btn-ghost btn-sm" style={{fontSize:11}} onClick={()=>edit(r)}>✏️</button><button className="btn btn-ghost btn-sm" style={{fontSize:11,color:"var(--red)"}} onClick={()=>del(r.id)}>🗑</button></div></td>
            </tr>
          ))}
        </tbody>
      </table></div>
    </div>
  );
}




// ══════════════════════════════════════════════════════════════
// ANALYTICS DASHBOARD COMPONENTS — Integrated from Analytics v3
// ══════════════════════════════════════════════════════════════


const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];

// Analytics constants used by dashboard components
const GRADES = ["OPC", "PPC", "ProMaxX"];
const REGIONS = ["Rajasthan", "Uttar Pradesh", "Haryana", "Madhya Pradesh"];
const ZONES = ["Zone A", "Zone B", "Zone C", "Zone D"];
const PLANTS = ["Aligarh Plant", "Morak Plant", "Depot"];
const PAY_TERMS = ["Advance", "Credit", "BG", "LC"];
const LEVELS = ["Sales Officer", "Area Sales Manager", "Regional Head", "Zonal Head", "Sales & Accounts", "Admin"];

// Analytics color palette

const G = {
  gold: "#C9A84C", goldDk: "#A07838", goldLt: "#F0D080", goldF: "#FFF9EE",
  grn: "#059669", red: "#DC2626", ora: "#EA580C", blue: "#2563EB",
  purple: "#7C3AED", ink: "#1C1917", muted: "#78716C", bg: "#FAFAF9",
  card: "#FFFFFF", border: "#E7E5E4", border2: "#D6D3D1",
};

const CHART_COLORS = [G.gold, G.blue, G.grn, G.purple, G.ora, G.red, "#0891B2", "#16A34A"];




/* ═══ HELPERS ════════════════════════════════════════════════════ */
const rnd = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;
const sum = (arr, k) => arr.reduce((s, r) => s + (r[k] || 0), 0);
const avg = (arr, k) => arr.length ? +(sum(arr, k) / arr.length).toFixed(1) : 0;
const pct = (n, d) => d ? Math.round(n / d * 100) : 0;

const AnalyticsBadge = ({ s }) => {
  const m = { Approved: "b-grn ✅", Pending: "b-gold 🕐", Rejected: "b-red ❌", Blocked: "b-ora 🔒" };
  const [cls, ic] = (m[s] || "b-muted •").split(" ");
  return <span className={`badge ${cls}`}>{ic} {s}</span>;
};

const TT = ({ active, payload, label, fmt }) => {
  if (!active || !payload || !payload.length) return null;
  return (
    <div style={{ background: G.card, border: `1.5px solid ${G.border}`, borderRadius: 8, padding: "10px 14px", fontSize: 12, boxShadow: "0 4px 12px #0002" }}>
      <div style={{ fontWeight: 700, marginBottom: 6, color: G.ink }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color, fontWeight: 600 }}>{p.name}: {fmt ? fmt(p.value) : p.value}</div>
      ))}
    </div>
  );
};



/* ═══ KPI CARD ════════════════════════════════════════════════════ */
const KPI = ({ icon, val, label, trend, trendTxt, color, sub }) => (
  <div className="kpi">
    <div className="kpi-accent" style={{ background: color || G.gold }} />
    <div className="kpi-icon">{icon}</div>
    <div className={`kpi-val${String(val).length > 8 ? " sm" : ""}`}>{val}</div>
    <div className="kpi-label">{label}</div>
    {trendTxt && <div className={`kpi-trend ${trend}`}>{trend === "up" ? "↑" : trend === "down" ? "↓" : "→"} {trendTxt}</div>}
    {sub && <div style={{ fontSize: 11, color: G.muted, marginTop: 2, fontWeight: 500 }}>{sub}</div>}
  </div>
);


const css = `
  @import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700;800&family=IBM+Plex+Sans:wght@400;500;600&display=swap');
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'IBM Plex Sans',sans-serif;background:${G.bg};color:${G.ink}}
  .app{display:flex;height:100vh;overflow:hidden}
  .sidebar{width:230px;background:#1C1917;flex-shrink:0;display:flex;flex-direction:column;overflow-y:auto}
  .sb-logo{padding:20px 16px 12px;border-bottom:1px solid #292524}
  .sb-brand{font-family:'Sora',sans-serif;font-weight:800;font-size:15px;color:${G.gold};letter-spacing:.3px}
  .sb-sub{font-size:10.5px;color:#A8A29E;font-weight:600;margin-top:2px;letter-spacing:.5px;text-transform:uppercase}
  .sb-section{padding:12px 10px 4px;font-size:9.5px;font-weight:700;color:#57534E;text-transform:uppercase;letter-spacing:.8px}
  .sb-item{display:flex;align-items:center;gap:9px;padding:8px 12px;border-radius:7px;margin:1px 6px;cursor:pointer;font-size:12.5px;font-weight:600;color:#A8A29E;transition:all .15s}
  .sb-item:hover{background:#292524;color:#E7E5E4}
  .sb-item.active{background:linear-gradient(90deg,${G.gold}22,${G.gold}11);color:${G.gold};border-left:3px solid ${G.gold};padding-left:9px}
  .sb-icon{font-size:14px;flex-shrink:0}
  .main{flex:1;overflow-y:auto;background:${G.bg}}
  .page{padding:24px 28px;min-height:100%}
  .ph{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:22px}
  .ph-ey{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:${G.gold};margin-bottom:3px}
  .ph-title{font-family:'Sora',sans-serif;font-weight:800;font-size:22px;color:${G.ink}}
  .ph-sub{font-size:12.5px;color:${G.muted};margin-top:3px;font-weight:500}
  .kpi-grid{display:grid;gap:14px;margin-bottom:20px}
  .kpi-grid-4{grid-template-columns:repeat(4,1fr)}
  .kpi-grid-5{grid-template-columns:repeat(5,1fr)}
  .kpi-grid-3{grid-template-columns:repeat(3,1fr)}
  .kpi{background:${G.card};border:1.5px solid ${G.border};border-radius:12px;padding:16px 18px;position:relative;overflow:hidden;cursor:default}
  .kpi-accent{position:absolute;top:0;left:0;right:0;height:3px}
  .kpi-icon{font-size:22px;margin-bottom:8px}
  .kpi-val{font-family:'Sora',sans-serif;font-weight:800;font-size:26px;color:${G.ink};line-height:1;margin-bottom:4px}
  .kpi-val.sm{font-size:18px}
  .kpi-label{font-size:11.5px;font-weight:600;color:${G.muted};text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px}
  .kpi-trend{font-size:11px;font-weight:700}
  .kpi-trend.up{color:${G.grn}} .kpi-trend.down{color:${G.red}} .kpi-trend.neu{color:${G.muted}}
  .card{background:${G.card};border:1.5px solid ${G.border};border-radius:12px;margin-bottom:16px}
  .card-hdr{display:flex;justify-content:space-between;align-items:center;padding:14px 18px;border-bottom:1.5px solid ${G.border}}
  .card-title{font-family:'Sora',sans-serif;font-weight:700;font-size:14px;color:${G.ink}}
  .card-sub{font-size:12px;color:${G.muted};margin-top:2px;font-weight:500}
  .card-body{padding:16px 18px}
  .grid-2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
  .grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px}
  .tw{overflow-x:auto}
  table{width:100%;border-collapse:collapse;font-size:12.5px}
  th{padding:10px 12px;text-align:left;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.4px;color:${G.muted};background:${G.bg};border-bottom:1.5px solid ${G.border}}
  td{padding:9px 12px;border-bottom:1px solid ${G.border};color:${G.ink};vertical-align:middle}
  tr:last-child td{border-bottom:none}
  tr:hover td{background:#FAFAF9}
  .badge{display:inline-flex;align-items:center;gap:4px;padding:3px 9px;border-radius:20px;font-size:11px;font-weight:700;border:1.5px solid}
  .b-grn{background:#ECFDF5;color:#065F46;border-color:#A7F3D0}
  .b-red{background:#FEF2F2;color:#991B1B;border-color:#FCA5A5}
  .b-ora{background:#FFF7ED;color:#9A3412;border-color:#FED7AA}
  .b-gold{background:${G.goldF};color:${G.goldDk};border-color:${G.goldLt}}
  .b-blue{background:#EFF6FF;color:#1E40AF;border-color:#BFDBFE}
  .b-muted{background:#F5F5F4;color:#57534E;border-color:#D6D3D1}
  .filter-bar{display:flex;gap:10px;flex-wrap:wrap;padding:12px 18px;border-bottom:1.5px solid ${G.border};align-items:center}
  .fsel{border:1.5px solid ${G.border2};border-radius:7px;padding:6px 10px;font-size:12px;font-family:inherit;background:${G.bg};color:${G.ink};font-weight:600;cursor:pointer;outline:none}
  .fsel:focus{border-color:${G.gold}}
  .sect-title{font-family:'Sora',sans-serif;font-weight:700;font-size:13px;color:${G.ink};padding:14px 18px 10px;border-bottom:1.5px solid ${G.border}}
  .funnel-step{display:flex;align-items:center;margin-bottom:8px;gap:12px}
  .funnel-bar{height:38px;border-radius:6px;display:flex;align-items:center;padding:0 14px;color:#fff;font-weight:700;font-size:13px;font-family:'Sora',sans-serif;transition:all .3s;min-width:60px}
  .tag{padding:2px 7px;border-radius:4px;font-size:10.5px;font-weight:700}
  .ncr-pos{color:${G.grn};font-family:'Sora',sans-serif;font-weight:800}
  .ncr-neg{color:${G.red};font-family:'Sora',sans-serif;font-weight:800}
  .aging-row{display:flex;align-items:center;gap:12px;padding:10px 18px;border-bottom:1px solid ${G.border}}
  .aging-bar-bg{flex:1;height:10px;background:${G.border};border-radius:5px;overflow:hidden}
  .aging-bar{height:100%;border-radius:5px;transition:width .4s}
  .stat-box{text-align:center;padding:14px;border-right:1px solid ${G.border}}
  .stat-box:last-child{border-right:none}
  .stat-v{font-family:'Sora',sans-serif;font-weight:800;font-size:22px}
  .stat-l{font-size:11px;font-weight:600;color:${G.muted};margin-top:3px;text-transform:uppercase;letter-spacing:.4px}
  .perf-row{display:flex;align-items:center;gap:14px;padding:12px 18px;border-bottom:1px solid ${G.border}}
  .perf-name{width:130px;font-weight:700;font-size:13px}
  .perf-role{font-size:10.5px;color:${G.muted};font-weight:600}
  .perf-bar-bg{flex:1;height:8px;background:${G.border};border-radius:4px;overflow:hidden}
  .perf-bar{height:100%;border-radius:4px;background:linear-gradient(90deg,${G.gold},${G.goldLt})}
  .tab-bar{display:flex;gap:0;border-bottom:2px solid ${G.border};margin-bottom:18px;overflow-x:auto}
  .tab{padding:9px 16px;font-size:12.5px;font-weight:600;cursor:pointer;color:${G.muted};border-bottom:2.5px solid transparent;margin-bottom:-2px;white-space:nowrap;transition:all .15s}
  .tab.act{color:${G.goldDk};border-color:${G.gold};font-weight:700}
  .tab:hover{color:${G.ink}}
  ::-webkit-scrollbar{width:5px;height:5px}
  ::-webkit-scrollbar-track{background:${G.bg}}
  ::-webkit-scrollbar-thumb{background:${G.border2};border-radius:3px}
`;

/* ════════════════════════════════════════════════════════════════
   DASHBOARD 1 — EXECUTIVE OVERVIEW
   ════════════════════════════════════════════════════════════════ */
function ExecDash({ req }) {
  const approved = req.filter(r => r.status === "Approved");
  const pending  = req.filter(r => r.status === "Pending");
  const rejected = req.filter(r => r.status === "Rejected");
  const blocked  = req.filter(r => r.status === "Blocked");
  const negNCR   = req.filter(r => r.ncrPmt < 0);
  const avgNCR   = avg(req, "ncrPmt");
  const totalVol = sum(approved, "qty");
  const totalRev = sum(approved, "orderPrice") * 1000;
  const avgTAT   = avg(req, "tat");
  const revImpact = (sum(approved, "orderPrice") * sum(approved, "qty") / 1000000).toFixed(1);

  const monthData = MONTHS.map(m => ({
    m,
    approved: req.filter(r => r.month === m && r.status === "Approved").length,
    pending:  req.filter(r => r.month === m && r.status === "Pending").length,
    rejected: req.filter(r => r.month === m && r.status === "Rejected").length,
    ncr: avg(req.filter(r => r.month === m), "ncrPmt"),
  }));

  const gradeVol = GRADES.map(g => ({
    name: g, vol: sum(approved.filter(r => r.grade === g), "qty"),
    count: approved.filter(r => r.grade === g).length,
  }));

  const regionData = REGIONS.map(r => ({
    name: r.replace("Uttar Pradesh", "UP"),
    count: req.filter(x => x.region === r).length,
    approved: approved.filter(x => x.region === r).length,
    vol: sum(approved.filter(x => x.region === r), "qty"),
  }));

  return (
    <div className="page">
      <div className="ph">
        <div>
          <div className="ph-ey">Analytics · Mangalam Cement · Birla Uttam</div>
          <div className="ph-title">Executive Overview</div>
          <div className="ph-sub">All-time summary · {req.length} total requests · Last updated: today</div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 28, color: G.gold }}>
            {pct(approved.length, req.length)}%
          </div>
          <div style={{ fontSize: 11, color: G.muted, fontWeight: 600, textTransform: "uppercase" }}>Overall Approval Rate</div>
        </div>
      </div>

      <div className="kpi-grid kpi-grid-5">
        <KPI icon="📋" val={req.length}          label="Total Requests"      trend="neu" trendTxt="All time"              color={G.gold}   />
        <KPI icon="✅" val={approved.length}      label="Approved"           trend="up"  trendTxt={`${pct(approved.length,req.length)}% rate`} color={G.grn}    />
        <KPI icon="🕐" val={pending.length}       label="Pending"            trend={pending.length>10?"down":"up"} trendTxt="In pipeline" color="#F59E0B"  />
        <KPI icon="❌" val={rejected.length}      label="Rejected"           trend="neu" trendTxt={`${pct(rejected.length,req.length)}% rate`} color={G.red}    />
        <KPI icon="🔒" val={blocked.length}       label="Blocked"            trend="neu" trendTxt="Need resolution"       color={G.ora}    />
      </div>
      <div className="kpi-grid kpi-grid-5">
        <KPI icon="📊" val={`₹${avgNCR}`}        label="Avg NCR/MT"         trend={avgNCR<0?"down":"up"} trendTxt="Per tonne"       color={avgNCR<0?G.red:G.grn} />
        <KPI icon="⚠️" val={negNCR.length}        label="Negative NCR Cases" trend="down" trendTxt={`${pct(negNCR.length,req.length)}% of total`} color={G.red}    />
        <KPI icon="📦" val={`${(totalVol/1000).toFixed(1)}K`} label="Approved Volume MT" trend="up" trendTxt="Metric tonnes"    color={G.purple} />
        <KPI icon="💰" val={`₹${revImpact}Cr`}   label="Revenue Impact"     trend="up"  trendTxt="Approved orders"       color={G.blue}   />
        <KPI icon="⏱️" val={`${avgTAT}d`}          label="Avg Approval TAT"   trend={avgTAT>5?"down":"up"} trendTxt="Days to approve" color="#0891B2"  />
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">Monthly Approval Trend</div><div className="card-sub">Requests by status per month</div></div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={monthData} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="m" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT />} />
                <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
                <Bar dataKey="approved" fill={G.grn}    name="Approved" radius={[3,3,0,0]} />
                <Bar dataKey="pending"  fill="#F59E0B"  name="Pending"  radius={[3,3,0,0]} />
                <Bar dataKey="rejected" fill={G.red}    name="Rejected" radius={[3,3,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="card">
          <div className="card-hdr"><div><div className="card-title">NCR Trend (Avg/MT)</div><div className="card-sub">Month-wise average NCR per tonne</div></div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={200}>
              <AreaChart data={monthData} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <defs>
                  <linearGradient id="ncrGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%"  stopColor={G.gold} stopOpacity={0.3} />
                    <stop offset="95%" stopColor={G.gold} stopOpacity={0.0} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="m" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT fmt={v => `₹${v}`} />} />
                <Area type="monotone" dataKey="ncr" stroke={G.gold} strokeWidth={2.5} fill="url(#ncrGrad)" name="Avg NCR/MT" dot={{ r: 3, fill: G.gold }} />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">Grade-wise Volume (MT)</div><div className="card-sub">Approved orders breakdown</div></div></div>
          <div className="card-body" style={{ display: "flex", alignItems: "center", gap: 20 }}>
            <ResponsiveContainer width="55%" height={170}>
              <PieChart>
                <Pie data={gradeVol} dataKey="vol" nameKey="name" cx="50%" cy="50%" outerRadius={65} innerRadius={35}>
                  {gradeVol.map((_, i) => <Cell key={i} fill={CHART_COLORS[i]} />)}
                </Pie>
                <Tooltip content={<TT fmt={v => `${v.toLocaleString()} MT`} />} />
              </PieChart>
            </ResponsiveContainer>
            <div style={{ flex: 1 }}>
              {gradeVol.map((g, i) => (
                <div key={g.name} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                  <div style={{ width: 10, height: 10, borderRadius: 2, background: CHART_COLORS[i], flexShrink: 0 }} />
                  <div style={{ flex: 1, fontSize: 12.5, fontWeight: 600 }}>{g.name}</div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: CHART_COLORS[i] }}>{g.vol.toLocaleString()} MT</div>
                </div>
              ))}
            </div>
          </div>
        </div>

        <div className="card">
          <div className="card-hdr"><div><div className="card-title">Region-wise Performance</div><div className="card-sub">Approved volume by region</div></div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={170}>
              <BarChart data={regionData} layout="vertical" margin={{ top: 0, right: 20, left: 20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} horizontal={false} />
                <XAxis type="number" tick={{ fontSize: 11 }} />
                <YAxis dataKey="name" type="category" tick={{ fontSize: 11 }} width={55} />
                <Tooltip content={<TT />} />
                <Bar dataKey="vol" name="Volume MT" radius={[0,4,4,0]}>
                  {regionData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ════════════════════════════════════════════════════════════════
   DASHBOARD 2 — APPROVAL WORKFLOW FUNNEL
   ════════════════════════════════════════════════════════════════ */
function WorkflowDash({ req }) {
  const levelColors = ["#F59E0B", "#3B82F6", "#8B5CF6", "#059669", "#0891B2", G.gold];
  const levelData = LEVELS.map((lvl, i) => {
    const atLevel  = req.filter(r => r.status === "Pending" && r.currentLevel === lvl);
    const passed   = req.filter(r => r.history && (r.history.some ? r.history.some(h => h.level === lvl) : false)).length;
    const rejected = req.filter(r => r.status === "Rejected" && r.currentLevel === lvl);
    const total    = req.filter(r => r.currentLevel === lvl || (r.history || []).some(h => h.level === lvl));
    return {
      name: lvl, short: ["SO", "ASM", "RH", "ZH", "S&A", "ADM"][i],
      pending: atLevel.length,
      color: levelColors[i],
      rejPct: pct(rejected.length, total.length || 1),
      avgTime: rnd(1, 5),
      processed: rnd(8, 25),
    };
  });

  const maxPending = Math.max(...levelData.map(l => l.pending), 1);

  const stageBar = LEVELS.map((lvl, i) => ({
    name: ["SO", "ASM", "RH", "ZH", "S&A", "ADM"][i],
    pending: levelData[i].pending,
    approved: rnd(5, 15),
    rejected: rnd(1, 4),
  }));

  return (
    <div className="page">
      <div className="ph">
        <div>
          <div className="ph-ey">Workflow Analytics</div>
          <div className="ph-title">Approval Workflow Dashboard</div>
          <div className="ph-sub">Live funnel · Stage-wise breakdown · TAT tracking</div>
        </div>
      </div>

      <div className="grid-2">
        {/* Funnel */}
        <div className="card">
          <div className="card-hdr"><div className="card-title">Approval Funnel — Pending Count</div></div>
          <div className="card-body">
            {levelData.map((lvl, i) => {
              const w = Math.max(30, Math.round((lvl.pending || 0.5) / maxPending * 85) + 15);
              return (
                <div key={lvl.name} className="funnel-step">
                  <div style={{ width: 36, fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 12, color: lvl.color, flexShrink: 0, textAlign: "right" }}>{lvl.short}</div>
                  <div style={{ flex: 1 }}>
                    <div className="funnel-bar" style={{ background: lvl.color, width: `${w}%` }}>
                      {lvl.pending} pending
                    </div>
                  </div>
                  <div style={{ width: 70, fontSize: 11, color: G.muted, fontWeight: 600, textAlign: "right", flexShrink: 0 }}>
                    ~{lvl.avgTime}d avg
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Stage stacked bar */}
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">Stage-wise Action Count</div><div className="card-sub">Processed, Approved, Rejected per level</div></div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={230}>
              <BarChart data={stageBar} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT />} />
                <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
                <Bar dataKey="approved" stackId="a" fill={G.grn}   name="Approved" radius={[0,0,0,0]} />
                <Bar dataKey="pending"  stackId="a" fill="#F59E0B" name="Pending"  />
                <Bar dataKey="rejected" stackId="a" fill={G.red}   name="Rejected" radius={[3,3,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* Level metrics table */}
      <div className="card">
        <div className="card-hdr"><div className="card-title">Level-wise Performance Metrics</div></div>
        <div className="tw">
          <table>
            <thead>
              <tr>
                <th>Level</th><th>Pending Now</th><th>Avg TAT (Days)</th><th>Rejection %</th><th>Requests Processed</th><th>Status</th>
              </tr>
            </thead>
            <tbody>
              {levelData.map((lvl, i) => (
                <tr key={lvl.name}>
                  <td><div style={{ display: "flex", alignItems: "center", gap: 8 }}><div style={{ width: 8, height: 8, borderRadius: 2, background: lvl.color }} /><span style={{ fontWeight: 700 }}>{lvl.name}</span></div></td>
                  <td><span style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 15, color: lvl.pending > 2 ? G.red : G.grn }}>{lvl.pending}</span></td>
                  <td><span style={{ fontWeight: 700, color: lvl.avgTime > 3 ? G.ora : G.grn }}>{lvl.avgTime}d</span></td>
                  <td><span className={`badge ${lvl.rejPct > 20 ? "b-red" : "b-grn"}`}>{lvl.rejPct}%</span></td>
                  <td style={{ fontWeight: 600 }}>{lvl.processed}</td>
                  <td><span className={`badge ${lvl.pending > 2 ? "b-ora" : "b-grn"}`}>{lvl.pending > 2 ? "⚠ Overloaded" : "✅ Normal"}</span></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ════════════════════════════════════════════════════════════════
   DASHBOARD 3 — NCR ANALYTICS
   ════════════════════════════════════════════════════════════════ */
function NCRDash({ req }) {
  const [fGrade, setFGrade] = useState("");
  const [fRegion, setFRegion] = useState("");
  const [fPlant, setFPlant] = useState("");
  const [fPay, setFPay] = useState("");

  const filtered = useMemo(() => req.filter(r =>
    (!fGrade  || r.grade  === fGrade)  &&
    (!fRegion || r.region === fRegion) &&
    (!fPlant  || r.plant  === fPlant)  &&
    (!fPay    || r.pay    === fPay)
  ), [req, fGrade, fRegion, fPlant, fPay]);

  const posNCR = filtered.filter(r => r.ncrPmt >= 0);
  const negNCR = filtered.filter(r => r.ncrPmt < 0);
  const avgN   = avg(filtered, "ncrPmt");
  const worst5 = [...filtered].sort((a, b) => a.ncrPmt - b.ncrPmt).slice(0, 5);
  const best5  = [...filtered].sort((a, b) => b.ncrPmt - a.ncrPmt).slice(0, 5);

  const gradeNCR = GRADES.map(g => ({ name: g, avg: avg(filtered.filter(r => r.grade === g), "ncrPmt"), count: filtered.filter(r => r.grade === g).length }));
  const plantNCR = PLANTS.map(p => ({ name: p.replace(" Plant", ""), avg: avg(filtered.filter(r => r.plant === p), "ncrPmt") }));
  const payNCR   = PAY_TERMS.map(p => ({ name: p, avg: avg(filtered.filter(r => r.pay === p), "ncrPmt") }));
  const monthNCR = MONTHS.map(m => ({ m, avg: avg(filtered.filter(r => r.month === m), "ncrPmt"), pos: filtered.filter(r => r.month === m && r.ncrPmt >= 0).length, neg: filtered.filter(r => r.month === m && r.ncrPmt < 0).length }));

  return (
    <div className="page">
      <div className="ph">
        <div>
          <div className="ph-ey">Business Critical</div>
          <div className="ph-title">NCR Analytics Dashboard</div>
          <div className="ph-sub">Net Contribution Ratio · {filtered.length} records · Primary margin intelligence</div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 28, color: avgN < 0 ? G.red : G.grn }}>₹{avgN}</div>
          <div style={{ fontSize: 11, color: G.muted, fontWeight: 600, textTransform: "uppercase" }}>Avg NCR / MT</div>
        </div>
      </div>

      {/* Filters */}
      <div className="card" style={{ marginBottom: 16 }}>
        <div className="filter-bar">
          <span style={{ fontSize: 12, fontWeight: 700, color: G.muted }}>Filter:</span>
          <select className="fsel" value={fGrade}  onChange={e => setFGrade(e.target.value)}>  <option value="">All Grades</option>  {GRADES.map(g    => <option key={g}>{g}</option>)}</select>
          <select className="fsel" value={fRegion} onChange={e => setFRegion(e.target.value)}><option value="">All Regions</option>{REGIONS.map(r   => <option key={r}>{r}</option>)}</select>
          <select className="fsel" value={fPlant}  onChange={e => setFPlant(e.target.value)}> <option value="">All Plants</option>  {PLANTS.map(p    => <option key={p}>{p}</option>)}</select>
          <select className="fsel" value={fPay}    onChange={e => setFPay(e.target.value)}>   <option value="">All Terms</option>   {PAY_TERMS.map(p => <option key={p}>{p}</option>)}</select>
          {(fGrade || fRegion || fPlant || fPay) && <button onClick={() => { setFGrade(""); setFRegion(""); setFPlant(""); setFPay(""); }} style={{ fontSize: 11, border: "none", background: "none", color: G.red, cursor: "pointer", fontWeight: 700 }}>✕ Clear</button>}
        </div>
      </div>

      <div className="kpi-grid kpi-grid-4">
        <KPI icon="📊" val={`₹${avgN}`}      label="Avg NCR/MT"      trend={avgN < 0 ? "down" : "up"} trendTxt="Overall"     color={avgN < 0 ? G.red : G.grn} />
        <KPI icon="✅" val={posNCR.length}    label="Positive NCR"    trend="up"  trendTxt={`${pct(posNCR.length,filtered.length)}%`} color={G.grn}  />
        <KPI icon="❌" val={negNCR.length}    label="Negative NCR"    trend="down" trendTxt={`${pct(negNCR.length,filtered.length)}%`} color={G.red}  />
        <KPI icon="⚠️" val={filtered.filter(r => r.ncrPmt < -300).length} label="Critical (<-300)" trend="down" trendTxt="Urgent review" color={G.ora} />
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div className="card-title">NCR Trend Month-wise</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={190}>
              <LineChart data={monthNCR} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="m" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT fmt={v => `₹${v}`} />} />
                <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
                <Line type="monotone" dataKey="avg" stroke={G.gold} strokeWidth={2.5} name="Avg NCR" dot={{ r: 3 }} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="card">
          <div className="card-hdr"><div className="card-title">Positive vs Negative NCR</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={190}>
              <BarChart data={monthNCR} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="m" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT />} />
                <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
                <Bar dataKey="pos" fill={G.grn}  name="Positive NCR" radius={[3,3,0,0]} />
                <Bar dataKey="neg" fill={G.red}  name="Negative NCR" radius={[3,3,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="grid-3">
        <div className="card">
          <div className="card-hdr"><div className="card-title">Grade-wise NCR</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={160}>
              <BarChart data={gradeNCR} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT fmt={v => `₹${v}`} />} />
                <Bar dataKey="avg" name="Avg NCR" radius={[4,4,0,0]}>
                  {gradeNCR.map((g, i) => <Cell key={i} fill={g.avg >= 0 ? G.grn : G.red} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
        <div className="card">
          <div className="card-hdr"><div className="card-title">Plant-wise NCR</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={160}>
              <BarChart data={plantNCR} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="name" tick={{ fontSize: 10 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT fmt={v => `₹${v}`} />} />
                <Bar dataKey="avg" name="Avg NCR" radius={[4,4,0,0]}>
                  {plantNCR.map((p, i) => <Cell key={i} fill={p.avg >= 0 ? G.blue : G.red} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
        <div className="card">
          <div className="card-hdr"><div className="card-title">Payment Terms Impact</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={160}>
              <BarChart data={payNCR} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT fmt={v => `₹${v}`} />} />
                <Bar dataKey="avg" name="Avg NCR" radius={[4,4,0,0]}>
                  {payNCR.map((p, i) => <Cell key={i} fill={CHART_COLORS[i]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">🔴 Top Negative NCR Deals</div><div className="card-sub">Worst margin requests — review urgently</div></div></div>
          <div className="tw">
            <table>
              <thead><tr><th>Request ID</th><th>Customer</th><th>Grade</th><th>Qty</th><th>NCR/MT</th><th>Total Loss</th></tr></thead>
              <tbody>
                {worst5.map(r => (
                  <tr key={r.id}>
                    <td style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: G.goldDk, fontSize: 12 }}>{r.id}</td>
                    <td style={{ fontWeight: 600 }}>{r.customer}</td>
                    <td><span className="badge b-muted">{r.grade}</span></td>
                    <td>{r.qty} MT</td>
                    <td><span className="ncr-neg">₹{r.ncrPmt}</span></td>
                    <td style={{ fontWeight: 700, color: G.red }}>₹{Math.abs(r.ncrPmt * r.qty / 20).toFixed(0)}K</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">🟢 Best Margin Requests</div><div className="card-sub">Top positive NCR deals</div></div></div>
          <div className="tw">
            <table>
              <thead><tr><th>Request ID</th><th>Customer</th><th>Grade</th><th>Qty</th><th>NCR/MT</th><th>Total Gain</th></tr></thead>
              <tbody>
                {best5.map(r => (
                  <tr key={r.id}>
                    <td style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: G.goldDk, fontSize: 12 }}>{r.id}</td>
                    <td style={{ fontWeight: 600 }}>{r.customer}</td>
                    <td><span className="badge b-muted">{r.grade}</span></td>
                    <td>{r.qty} MT</td>
                    <td><span className="ncr-pos">₹{r.ncrPmt}</span></td>
                    <td style={{ fontWeight: 700, color: G.grn }}>₹{Math.abs(r.ncrPmt * r.qty / 20).toFixed(0)}K</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ════════════════════════════════════════════════════════════════
   DASHBOARD 4 — TEAM PERFORMANCE
   ════════════════════════════════════════════════════════════════ */
function PerfDash({ req }) {
  const [roleTab, setRoleTab] = useState("SO");
  const ROLE_MAP = { SO: "Sales Officer", ASM: "Area Sales Manager", RH: "Regional Head", ZH: "Zonal Head" };

  const soData = ["SO001", "SO002", "SO003", "SO004"].map((id, i) => {
    const mine = req.filter(r => r.soId === `SO00${i + 1}`);
    const app  = mine.filter(r => r.status === "Approved");
    return {
      id, name: ["Amit Sharma", "Priya Singh", "Ravi Gupta", "Neha Verma"][i],
      role: "Sales Officer", created: mine.length,
      approved: app.length, rejected: mine.filter(r => r.status === "Rejected").length,
      appPct: pct(app.length, mine.length || 1),
      avgNCR: avg(mine, "ncrPmt"),
      vol: sum(app, "qty"),
      avgTAT: rnd(2, 8),
    };
  });

  const radarData = soData.map(s => ({
    name: s.name.split(" ")[0],
    Volume: Math.min(100, Math.round(s.vol / 30)),
    Approval: s.appPct,
    NCR: Math.max(0, Math.min(100, 50 + s.avgNCR)),
    Speed: Math.max(0, 100 - s.avgTAT * 10),
    Count: Math.min(100, s.created * 5),
  }));

  return (
    <div className="page">
      <div className="ph">
        <div>
          <div className="ph-ey">HR Analytics</div>
          <div className="ph-title">Team Performance Dashboard</div>
          <div className="ph-sub">Individual tracking · Approval efficiency · Volume managed</div>
        </div>
      </div>

      <div className="tab-bar">
        {Object.keys(ROLE_MAP).map(k => (
          <div key={k} className={`tab ${roleTab === k ? "act" : ""}`} onClick={() => setRoleTab(k)}>{k} — {ROLE_MAP[k]}</div>
        ))}
      </div>

      <div className="kpi-grid kpi-grid-4" style={{ marginBottom: 16 }}>
        {soData.map(s => (
          <div key={s.id} className="kpi">
            <div className="kpi-accent" style={{ background: G.gold }} />
            <div style={{ display: "flex", align: "center", gap: 10, marginBottom: 10 }}>
              <div style={{ width: 36, height: 36, borderRadius: "50%", background: `linear-gradient(135deg,${G.gold},${G.goldDk})`, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 14 }}>
                {s.name[0]}
              </div>
              <div>
                <div style={{ fontWeight: 700, fontSize: 13 }}>{s.name}</div>
                <div style={{ fontSize: 10.5, color: G.muted, fontWeight: 600 }}>Sales Officer</div>
              </div>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6, marginTop: 4 }}>
              <div><div style={{ fontSize: 10, color: G.muted, fontWeight: 600 }}>REQUESTS</div><div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 16 }}>{s.created}</div></div>
              <div><div style={{ fontSize: 10, color: G.muted, fontWeight: 600 }}>APPROVED</div><div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 16, color: G.grn }}>{s.appPct}%</div></div>
              <div><div style={{ fontSize: 10, color: G.muted, fontWeight: 600 }}>AVG NCR</div><div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 14, color: s.avgNCR < 0 ? G.red : G.grn }}>₹{s.avgNCR}</div></div>
              <div><div style={{ fontSize: 10, color: G.muted, fontWeight: 600 }}>VOLUME</div><div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 14 }}>{s.vol.toLocaleString()}</div></div>
            </div>
            <div style={{ marginTop: 10 }}>
              <div style={{ fontSize: 10, color: G.muted, fontWeight: 600, marginBottom: 4 }}>APPROVAL RATE</div>
              <div style={{ background: G.border, borderRadius: 4, height: 6, overflow: "hidden" }}>
                <div style={{ height: "100%", width: `${s.appPct}%`, background: `linear-gradient(90deg,${G.gold},${G.goldLt})`, borderRadius: 4 }} />
              </div>
            </div>
          </div>
        ))}
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div className="card-title">Performance Radar — Sales Officers</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={230}>
              <RadarChart data={radarData}>
                <PolarGrid stroke={G.border} />
                <PolarAngleAxis dataKey="name" tick={{ fontSize: 11 }} />
                {soData.map((s, i) => (
                  <Radar key={s.id} name={s.name.split(" ")[0]} dataKey={s.name.split(" ")[0]}
                    stroke={CHART_COLORS[i]} fill={CHART_COLORS[i]} fillOpacity={0.1} />
                ))}
              </RadarChart>
            </ResponsiveContainer>
          </div>
        </div>
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">Leaderboard</div><div className="card-sub">Ranked by approval rate</div></div></div>
          <div style={{ padding: 0 }}>
            {[...soData].sort((a, b) => b.appPct - a.appPct).map((s, i) => (
              <div key={s.id} className="perf-row">
                <div style={{ width: 22, height: 22, borderRadius: "50%", background: [G.gold, G.muted, "#CD7F32", G.border2][i], color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 11, flexShrink: 0 }}>{i + 1}</div>
                <div className="perf-name">{s.name}<div className="perf-role">{s.created} requests</div></div>
                <div style={{ flex: 1 }}>
                  <div className="perf-bar-bg"><div className="perf-bar" style={{ width: `${s.appPct}%` }} /></div>
                </div>
                <div style={{ width: 50, textAlign: "right", fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 15, color: G.grn }}>{s.appPct}%</div>
                <div style={{ width: 60, textAlign: "right", fontSize: 11.5, fontWeight: 700, color: s.avgNCR < 0 ? G.red : G.grn }}>₹{s.avgNCR}</div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ════════════════════════════════════════════════════════════════
   DASHBOARD 5 — SALES vs APPROVAL
   ════════════════════════════════════════════════════════════════ */
function SalesDash({ req }) {
  const approved = req.filter(r => r.status === "Approved");
  const utilized = approved.map(r => ({
    ...r,
    utilPct: r.qty > 0 ? pct(r.invoicedQty, r.qty) : 0,
    gap: r.qty - r.invoicedQty,
  }));

  const totalApprQty  = sum(approved, "qty");
  const totalSoldQty  = sum(approved, "invoicedQty");
  const totalGapQty   = totalApprQty - totalSoldQty;
  const overallUtilPct = pct(totalSoldQty, totalApprQty);

  const monthUtil = MONTHS.map(m => {
    const ma = approved.filter(r => r.month === m);
    const ap = sum(ma, "qty");
    const so = sum(ma, "invoicedQty");
    return { m, approved: ap, sold: so, gap: ap - so, utilPct: pct(so, ap) };
  });

  const gradeUtil = GRADES.map(g => {
    const ga = approved.filter(r => r.grade === g);
    const ap = sum(ga, "qty");
    const so = sum(ga, "invoicedQty");
    return { name: g, approved: ap, sold: so, utilPct: pct(so, ap) };
  });

  const lowUtil = utilized.filter(r => r.utilPct < 50 && r.qty > 0).slice(0, 5);

  return (
    <div className="page">
      <div className="ph">
        <div>
          <div className="ph-ey">Sales Intelligence</div>
          <div className="ph-title">Sales vs Approval Dashboard</div>
          <div className="ph-sub">Approved quantity utilization · Invoice tracking · Gap analysis</div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 28, color: overallUtilPct > 70 ? G.grn : G.ora }}>{overallUtilPct}%</div>
          <div style={{ fontSize: 11, color: G.muted, fontWeight: 600, textTransform: "uppercase" }}>Overall Utilization</div>
        </div>
      </div>

      <div className="kpi-grid kpi-grid-4">
        <KPI icon="📦" val={`${(totalApprQty/1000).toFixed(1)}K`} label="Total Approved Qty (MT)" trend="up" trendTxt="Active approvals"  color={G.gold}   />
        <KPI icon="✅" val={`${(totalSoldQty/1000).toFixed(1)}K`} label="Total Invoiced Qty (MT)"  trend="up" trendTxt={`${overallUtilPct}% utilized`} color={G.grn}    />
        <KPI icon="⏳" val={`${(totalGapQty/1000).toFixed(1)}K`}  label="Pending Qty (Gap)"        trend="neu" trendTxt="Yet to invoice"    color={G.ora}    />
        <KPI icon="📋" val={lowUtil.length}                        label="Low Utilization (<50%)"   trend="down" trendTxt="Need follow-up"   color={G.red}    />
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div className="card-title">Monthly Approved vs Sold Qty</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={monthUtil} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="m" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT />} />
                <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
                <Bar dataKey="approved" fill={G.blue}  name="Approved Qty" radius={[3,3,0,0]} />
                <Bar dataKey="sold"     fill={G.grn}   name="Sold Qty"     radius={[3,3,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="card">
          <div className="card-hdr"><div className="card-title">Grade-wise Utilization %</div></div>
          <div className="card-body">
            {gradeUtil.map((g, i) => (
              <div key={g.name} style={{ marginBottom: 16 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 5 }}>
                  <span style={{ fontWeight: 700, fontSize: 13 }}>{g.name}</span>
                  <span style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 13, color: g.utilPct > 70 ? G.grn : G.ora }}>{g.utilPct}%</span>
                </div>
                <div style={{ background: G.border, borderRadius: 5, height: 10, overflow: "hidden" }}>
                  <div style={{ height: "100%", width: `${g.utilPct}%`, background: CHART_COLORS[i], borderRadius: 5, transition: "width .4s" }} />
                </div>
                <div style={{ fontSize: 11, color: G.muted, marginTop: 3 }}>{g.sold.toLocaleString()} / {g.approved.toLocaleString()} MT sold</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="card">
        <div className="card-hdr"><div><div className="card-title">⚠ Low Utilization Orders (&lt;50%)</div><div className="card-sub">Approved but not invoiced — needs follow-up</div></div></div>
        <div className="tw">
          <table>
            <thead><tr><th>Request ID</th><th>Customer</th><th>Grade</th><th>Region</th><th>Approved Qty</th><th>Invoiced Qty</th><th>Gap</th><th>Utilization</th></tr></thead>
            <tbody>
              {lowUtil.map(r => (
                <tr key={r.id}>
                  <td style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: G.goldDk, fontSize: 12 }}>{r.id}</td>
                  <td style={{ fontWeight: 600 }}>{r.customer}</td>
                  <td><span className="badge b-muted">{r.grade}</span></td>
                  <td style={{ fontSize: 12 }}>{r.region}</td>
                  <td style={{ fontWeight: 700 }}>{r.qty} MT</td>
                  <td style={{ fontWeight: 600, color: G.grn }}>{r.invoicedQty} MT</td>
                  <td style={{ fontWeight: 700, color: G.red }}>{r.gap} MT</td>
                  <td>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <div style={{ width: 60, height: 6, background: G.border, borderRadius: 3, overflow: "hidden" }}>
                        <div style={{ height: "100%", width: `${r.utilPct}%`, background: r.utilPct < 30 ? G.red : G.ora, borderRadius: 3 }} />
                      </div>
                      <span style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 12, color: G.red }}>{r.utilPct}%</span>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ════════════════════════════════════════════════════════════════
   DASHBOARD 6 — CUSTOMER INTELLIGENCE
   ════════════════════════════════════════════════════════════════ */
function CustomerDash({ req }) {
  const customers = ["C001", "C002", "C003", "C004", "C005", "C006", "C007", "C008"].map((id, i) => {
    const mine = req.filter(r => r.customer === `Customer ${id}`);
    const app  = mine.filter(r => r.status === "Approved");
    const blk  = mine.filter(r => r.status === "Blocked");
    return {
      id, name: ["Raj Construction", "Sharma Builders", "Gupta Infra", "Mehta Projects", "Patel Constructions", "Singh Infra", "Kumar Developers", "Verma Traders"][i],
      requests: mine.length, approved: app.length, blocked: blk.length,
      vol: sum(app, "qty"), avgNCR: avg(mine, "ncrPmt"),
      avgDiscount: avg(mine, "difference"),
      region: REGIONS[i % 4],
    };
  });

  const topVol      = [...customers].sort((a, b) => b.vol - a.vol).slice(0, 5);
  const topDisc     = [...customers].sort((a, b) => a.avgDiscount - b.avgDiscount).slice(0, 5);
  const topNegNCR   = [...customers].sort((a, b) => a.avgNCR - b.avgNCR).slice(0, 5);
  const mostFreq    = [...customers].sort((a, b) => b.requests - a.requests).slice(0, 5);

  const custPie = customers.slice(0, 5).map(c => ({ name: c.name.split(" ")[0], vol: c.vol }));

  return (
    <div className="page">
      <div className="ph">
        <div>
          <div className="ph-ey">CRM Intelligence</div>
          <div className="ph-title">Customer Intelligence Dashboard</div>
          <div className="ph-sub">Risk · Volume · Discount · NCR — per customer</div>
        </div>
      </div>

      <div className="kpi-grid kpi-grid-4">
        <KPI icon="🏢" val={customers.length}                     label="Total Customers"      trend="neu" trendTxt="In system"      color={G.gold}   />
        <KPI icon="🔒" val={customers.filter(c=>c.blocked>0).length} label="Blocked Customers"  trend="down" trendTxt="Need review"  color={G.red}    />
        <KPI icon="📊" val={customers.filter(c=>c.avgNCR<0).length}  label="Neg NCR Customers"  trend="down" trendTxt="Loss accounts" color={G.ora}    />
        <KPI icon="🌟" val={customers.filter(c=>c.vol>500).length}   label="High Volume Cust."  trend="up"  trendTxt=">500 MT"        color={G.grn}    />
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div className="card-title">Customer Volume Share</div></div>
          <div className="card-body" style={{ display: "flex", alignItems: "center", gap: 16 }}>
            <ResponsiveContainer width="50%" height={170}>
              <PieChart>
                <Pie data={custPie} dataKey="vol" nameKey="name" cx="50%" cy="50%" outerRadius={60} innerRadius={30}>
                  {custPie.map((_, i) => <Cell key={i} fill={CHART_COLORS[i]} />)}
                </Pie>
                <Tooltip content={<TT fmt={v => `${v} MT`} />} />
              </PieChart>
            </ResponsiveContainer>
            <div style={{ flex: 1 }}>
              {custPie.map((c, i) => (
                <div key={c.name} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                  <div style={{ width: 8, height: 8, borderRadius: 2, background: CHART_COLORS[i] }} />
                  <span style={{ fontSize: 12, fontWeight: 600, flex: 1 }}>{c.name}</span>
                  <span style={{ fontSize: 12, fontWeight: 700 }}>{c.vol} MT</span>
                </div>
              ))}
            </div>
          </div>
        </div>
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">Top Volume Customers</div><div className="card-sub">By approved MT</div></div></div>
          <div style={{ padding: 0 }}>
            {topVol.map((c, i) => (
              <div key={c.id} className="perf-row">
                <div style={{ width: 20, fontWeight: 800, color: G.gold, fontSize: 13 }}>{i + 1}</div>
                <div className="perf-name">{c.name}<div className="perf-role">{c.region}</div></div>
                <div style={{ flex: 1 }}>
                  <div className="perf-bar-bg"><div className="perf-bar" style={{ width: `${pct(c.vol, topVol[0].vol)}%` }} /></div>
                </div>
                <div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: G.blue, minWidth: 70, textAlign: "right" }}>{c.vol.toLocaleString()} MT</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">🔴 High Discount Customers</div><div className="card-sub">Most negative difference from trade price</div></div></div>
          <div className="tw">
            <table>
              <thead><tr><th>Customer</th><th>Avg Discount</th><th>Requests</th><th>Avg NCR</th></tr></thead>
              <tbody>
                {topDisc.map(c => (
                  <tr key={c.id}>
                    <td style={{ fontWeight: 700 }}>{c.name}</td>
                    <td style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: G.red }}>₹{c.avgDiscount.toFixed(0)}</td>
                    <td>{c.requests}</td>
                    <td><span className={c.avgNCR < 0 ? "ncr-neg" : "ncr-pos"}>₹{c.avgNCR}</span></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">Most Frequently Approved</div><div className="card-sub">Repeat business customers</div></div></div>
          <div className="tw">
            <table>
              <thead><tr><th>Customer</th><th>Requests</th><th>Approved</th><th>Vol (MT)</th></tr></thead>
              <tbody>
                {mostFreq.map(c => (
                  <tr key={c.id}>
                    <td style={{ fontWeight: 700 }}>{c.name}</td>
                    <td style={{ fontWeight: 600 }}>{c.requests}</td>
                    <td><span className="badge b-grn">✅ {c.approved}</span></td>
                    <td style={{ fontFamily: "'Sora',sans-serif", fontWeight: 700, color: G.blue }}>{c.vol}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ════════════════════════════════════════════════════════════════
   DASHBOARD 7 — FREIGHT & COST
   ════════════════════════════════════════════════════════════════ */
function FreightDash({ req }) {
  const approved = req.filter(r => r.status === "Approved");

  const costBreakdown = [
    { name: "Primary Freight", avg: avg(req, "primaryFreight"),   color: G.gold,   key: "primaryFreight" },
    { name: "Secondary Freight", avg: avg(req, "secondaryFreight"), color: G.blue, key: "secondaryFreight" },
    { name: "Cost of Production", avg: avg(req, "cop"),            color: G.purple, key: "cop" },
    { name: "Packing",           avg: avg(req, "packing"),          color: G.grn,   key: "packing" },
    { name: "OP Commission",     avg: avg(req, "opCommission"),     color: G.ora,   key: "opCommission" },
  ];
  const totalCostAvg = costBreakdown.reduce((s, c) => s + c.avg, 0);

  const plantFreight = PLANTS.map(p => ({
    name: p.replace(" Plant", ""),
    pf: avg(req.filter(r => r.plant === p), "primaryFreight"),
    sf: avg(req.filter(r => r.plant === p), "secondaryFreight"),
    cop: avg(req.filter(r => r.plant === p), "cop"),
  }));

  const pieData = costBreakdown.map(c => ({ name: c.name.split(" ")[0], value: c.avg, color: c.color }));

  return (
    <div className="page">
      <div className="ph">
        <div>
          <div className="ph-ey">Cost Analytics</div>
          <div className="ph-title">Freight & Cost Analytics</div>
          <div className="ph-sub">Cost component breakdown · Leakage analysis · Route comparison</div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 26, color: G.ink }}>₹{totalCostAvg.toFixed(0)}</div>
          <div style={{ fontSize: 11, color: G.muted, fontWeight: 600, textTransform: "uppercase" }}>Avg Total Cost/MT</div>
        </div>
      </div>

      <div className="kpi-grid kpi-grid-5">
        {costBreakdown.map(c => (
          <KPI key={c.name} icon="💸" val={`₹${c.avg}`} label={c.name} trend="neu" trendTxt={`${pct(c.avg, totalCostAvg)}% of total`} color={c.color} />
        ))}
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div className="card-title">Cost Component Pie</div></div>
          <div className="card-body" style={{ display: "flex", alignItems: "center", gap: 16 }}>
            <ResponsiveContainer width="55%" height={180}>
              <PieChart>
                <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={70} innerRadius={40}>
                  {pieData.map((p, i) => <Cell key={i} fill={p.color} />)}
                </Pie>
                <Tooltip content={<TT fmt={v => `₹${v.toFixed(1)}`} />} />
              </PieChart>
            </ResponsiveContainer>
            <div style={{ flex: 1 }}>
              {costBreakdown.map(c => (
                <div key={c.name} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                  <div style={{ width: 8, height: 8, borderRadius: 2, background: c.color, flexShrink: 0 }} />
                  <div style={{ flex: 1, fontSize: 12, fontWeight: 600 }}>{c.name}</div>
                  <div style={{ fontSize: 12, fontWeight: 700 }}>₹{c.avg}</div>
                  <div style={{ fontSize: 11, color: G.muted, width: 30, textAlign: "right" }}>{pct(c.avg, totalCostAvg)}%</div>
                </div>
              ))}
            </div>
          </div>
        </div>

        <div className="card">
          <div className="card-hdr"><div className="card-title">Plant-wise Cost Comparison</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={plantFreight} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT fmt={v => `₹${v.toFixed(0)}`} />} />
                <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
                <Bar dataKey="pf"  fill={G.gold}   name="Primary Freight"   radius={[3,3,0,0]} />
                <Bar dataKey="sf"  fill={G.blue}   name="Secondary Freight" radius={[3,3,0,0]} />
                <Bar dataKey="cop" fill={G.purple} name="COP"               radius={[3,3,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="card">
        <div className="card-hdr"><div><div className="card-title">Cost Leakage Analysis</div><div className="card-sub">Which component drives highest NCR loss</div></div></div>
        <div className="card-body">
          {costBreakdown.map(c => {
            const contribPct = pct(c.avg, totalCostAvg);
            return (
              <div key={c.name} style={{ marginBottom: 14 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                  <span style={{ fontWeight: 700, fontSize: 13 }}>{c.name}</span>
                  <div style={{ display: "flex", gap: 12 }}>
                    <span style={{ fontSize: 12.5, fontWeight: 700 }}>₹{c.avg}/MT</span>
                    <span style={{ fontSize: 12, fontWeight: 700, color: c.color }}>{contribPct}% of cost</span>
                  </div>
                </div>
                <div style={{ background: G.border, borderRadius: 5, height: 10, overflow: "hidden" }}>
                  <div style={{ height: "100%", width: `${contribPct}%`, background: c.color, borderRadius: 5 }} />
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

/* ════════════════════════════════════════════════════════════════
   DASHBOARD 8 — APPROVAL AGING
   ════════════════════════════════════════════════════════════════ */
function AgingDash({ req }) {
  const pending = req.filter(r => r.status === "Pending");
  const ageGroups = [
    { label: "0–1 Day",  min: 0, max: 1,  color: G.grn },
    { label: "2–3 Days", min: 2, max: 3,  color: G.gold },
    { label: "4–7 Days", min: 4, max: 7,  color: G.ora },
    { label: "> 7 Days", min: 8, max: 99, color: G.red },
  ];

  const aged = ageGroups.map(g => ({
    ...g,
    count: pending.filter(r => r.tat >= g.min && r.tat <= g.max).length,
    list:  pending.filter(r => r.tat >= g.min && r.tat <= g.max).slice(0, 3),
  }));
  const maxCount = Math.max(...aged.map(a => a.count), 1);

  const levelAging = LEVELS.slice(0, 5).map(lvl => ({
    name: lvl, stuck: pending.filter(r => r.currentLevel === lvl && r.tat > 3).length,
    avgDays: avg(pending.filter(r => r.currentLevel === lvl), "tat"),
  }));

  const agingBar = aged.map(a => ({ name: a.label, count: a.count }));

  return (
    <div className="page">
      <div className="ph">
        <div>
          <div className="ph-ey">Operational</div>
          <div className="ph-title">Approval Aging Dashboard</div>
          <div className="ph-sub">Pending requests · Escalation alerts · TAT analysis</div>
        </div>
        <div style={{ padding: "8px 16px", background: pending.filter(r => r.tat > 7).length > 0 ? "#FEF2F2" : "#ECFDF5", borderRadius: 10, border: `1.5px solid ${pending.filter(r => r.tat > 7).length > 0 ? "#FCA5A5" : "#A7F3D0"}` }}>
          <div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 22, color: pending.filter(r => r.tat > 7).length > 0 ? G.red : G.grn }}>
            {pending.filter(r => r.tat > 7).length}
          </div>
          <div style={{ fontSize: 11, fontWeight: 600, color: G.muted, textTransform: "uppercase" }}>Critical (&gt;7 days)</div>
        </div>
      </div>

      <div className="kpi-grid kpi-grid-4">
        {aged.map(a => <KPI key={a.label} icon="📋" val={a.count} label={a.label} trend={a.min > 3 ? "down" : "up"} trendTxt="Pending requests" color={a.color} />)}
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div className="card-title">Aging Bucket Chart</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={agingBar} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT />} />
                <Bar dataKey="count" name="Pending Count" radius={[5,5,0,0]}>
                  {agingBar.map((_, i) => <Cell key={i} fill={aged[i].color} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">Escalation Risk by Level</div><div className="card-sub">Requests stuck &gt;3 days per level</div></div></div>
          <div style={{ padding: 0 }}>
            {levelAging.map(l => (
              <div key={l.name} className="aging-row">
                <div style={{ width: 110, fontSize: 12, fontWeight: 700 }}>{l.name.replace("Area Sales Manager", "ASM").replace("Regional Head", "RH").replace("Zonal Head", "ZH").replace("Sales & Accounts", "S&A")}</div>
                <div className="aging-bar-bg">
                  <div className="aging-bar" style={{ width: `${pct(l.stuck, Math.max(...levelAging.map(x => x.stuck), 1))}%`, background: l.stuck > 2 ? G.red : l.stuck > 0 ? G.ora : G.grn }} />
                </div>
                <div style={{ width: 70, fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 14, color: l.stuck > 2 ? G.red : l.stuck > 0 ? G.ora : G.grn, textAlign: "right" }}>{l.stuck} stuck</div>
                <span className={`badge ${l.avgDays > 5 ? "b-red" : l.avgDays > 2 ? "b-ora" : "b-grn"}`}>{l.avgDays.toFixed(1)}d avg</span>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="card">
        <div className="card-hdr"><div><div className="card-title">🚨 Critical Pending — &gt;7 Days</div><div className="card-sub">Immediate escalation required</div></div></div>
        <div className="tw">
          <table>
            <thead><tr><th>Request ID</th><th>Customer</th><th>Grade</th><th>Region</th><th>Current Level</th><th>Days Pending</th><th>Priority</th></tr></thead>
            <tbody>
              {pending.filter(r => r.tat > 5).slice(0, 8).map(r => (
                <tr key={r.id}>
                  <td style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: G.goldDk, fontSize: 12 }}>{r.id}</td>
                  <td style={{ fontWeight: 600 }}>{r.customer}</td>
                  <td><span className="badge b-muted">{r.grade}</span></td>
                  <td style={{ fontSize: 12 }}>{r.region}</td>
                  <td style={{ fontWeight: 600, fontSize: 12 }}>{r.currentLevel}</td>
                  <td><span style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: r.tat > 7 ? G.red : G.ora }}>{r.tat}d</span></td>
                  <td><span className={`badge ${r.tat > 7 ? "b-red" : "b-ora"}`}>{r.tat > 7 ? "🚨 Critical" : "⚠ High"}</span></td>
                </tr>
              ))}
              {pending.filter(r => r.tat > 5).length === 0 && (
                <tr><td colSpan={7} style={{ textAlign: "center", padding: 30, color: G.grn, fontWeight: 700 }}>✅ No critical aging requests</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ════════════════════════════════════════════════════════════════
   DASHBOARD 9 — TRADE PRICE ANALYSIS
   ════════════════════════════════════════════════════════════════ */
function TradeDash({ req }) {
  const monthDisc = MONTHS.map(m => ({
    m, avgTrade: avg(req.filter(r => r.month === m), "tradePrice"),
    avgOrder: avg(req.filter(r => r.month === m), "orderPrice"),
    avgDiff:  avg(req.filter(r => r.month === m), "difference"),
  }));

  const gradeDisc = GRADES.map(g => ({
    name: g,
    trade: avg(req.filter(r => r.grade === g), "tradePrice"),
    order: avg(req.filter(r => r.grade === g), "orderPrice"),
    diff:  avg(req.filter(r => r.grade === g), "difference"),
  }));

  const regionDisc = REGIONS.map(r => ({
    name: r.replace("Uttar Pradesh", "UP"),
    diff: avg(req.filter(x => x.region === r), "difference"),
    count: req.filter(x => x.region === r).length,
  }));

  const avgDiff = avg(req, "difference");

  return (
    <div className="page">
      <div className="ph">
        <div>
          <div className="ph-ey">Pricing Intelligence</div>
          <div className="ph-title">Trade Price vs Non-Trade Analysis</div>
          <div className="ph-sub">Discount tracking · Market competitiveness · Price realization</div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 28, color: avgDiff < 0 ? G.red : G.grn }}>₹{avgDiff.toFixed(0)}</div>
          <div style={{ fontSize: 11, color: G.muted, fontWeight: 600, textTransform: "uppercase" }}>Avg Discount/MT</div>
        </div>
      </div>

      <div className="kpi-grid kpi-grid-4">
        <KPI icon="💰" val={`₹${avg(req,"tradePrice").toFixed(0)}`} label="Avg Trade Price"  trend="neu" trendTxt="Reference price"   color={G.gold}   />
        <KPI icon="📋" val={`₹${avg(req,"orderPrice").toFixed(0)}`} label="Avg Order Price"  trend="neu" trendTxt="Actual deal price"  color={G.blue}   />
        <KPI icon="📉" val={`₹${Math.abs(avgDiff).toFixed(0)}`}    label="Avg Discount"      trend="down" trendTxt="Below trade price" color={G.red}    />
        <KPI icon="🎯" val={`${pct(req.filter(r=>r.difference>=0).length,req.length)}%`} label="Above Trade Price" trend="up" trendTxt="Premium deals" color={G.grn}    />
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div className="card-title">Trade vs Order Price Trend</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={200}>
              <LineChart data={monthDisc} margin={{ top: 0, right: 0, left: -10, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="m" tick={{ fontSize: 11 }} />
                <YAxis domain={["auto", "auto"]} tick={{ fontSize: 11 }} />
                <Tooltip content={<TT fmt={v => `₹${v.toFixed(0)}`} />} />
                <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
                <Line type="monotone" dataKey="avgTrade" stroke={G.gold}  strokeWidth={2} name="Trade Price" dot={{ r: 3 }} strokeDasharray="5 3" />
                <Line type="monotone" dataKey="avgOrder" stroke={G.blue}  strokeWidth={2} name="Order Price" dot={{ r: 3 }} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>
        <div className="card">
          <div className="card-hdr"><div className="card-title">Monthly Discount Trend</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={200}>
              <AreaChart data={monthDisc} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <defs>
                  <linearGradient id="discGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%"  stopColor={G.red} stopOpacity={0.25} />
                    <stop offset="95%" stopColor={G.red} stopOpacity={0.0} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="m" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT fmt={v => `₹${v.toFixed(0)}`} />} />
                <Area type="monotone" dataKey="avgDiff" stroke={G.red} strokeWidth={2.5} fill="url(#discGrad)" name="Avg Discount" dot={{ r: 3, fill: G.red }} />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div className="card-title">Grade-wise Price Comparison</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={180}>
              <BarChart data={gradeDisc} margin={{ top: 0, right: 0, left: -10, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT fmt={v => `₹${v.toFixed(0)}`} />} />
                <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
                <Bar dataKey="trade" fill={G.gold} name="Trade Price" radius={[3,3,0,0]} />
                <Bar dataKey="order" fill={G.blue} name="Order Price" radius={[3,3,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">Region Competitiveness Index</div><div className="card-sub">Avg discount per region (negative = below trade)</div></div></div>
          <div className="card-body">
            {regionDisc.map((r, i) => (
              <div key={r.name} style={{ marginBottom: 14 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                  <span style={{ fontWeight: 700, fontSize: 13 }}>{r.name}</span>
                  <span style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: r.diff < -15 ? G.red : r.diff < 0 ? G.ora : G.grn }}>₹{r.diff.toFixed(0)}/MT</span>
                </div>
                <div style={{ background: G.border, borderRadius: 5, height: 10, overflow: "hidden" }}>
                  <div style={{ height: "100%", width: `${Math.min(100, Math.abs(r.diff) * 1.5)}%`, background: r.diff < -15 ? G.red : r.diff < 0 ? G.ora : G.grn, borderRadius: 5 }} />
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ════════════════════════════════════════════════════════════════
   DASHBOARD 10 — GEOGRAPHIC
   ════════════════════════════════════════════════════════════════ */
function GeoDash({ req }) {
  const districts = ["Jaipur", "Jodhpur", "Kota", "Alwar", "Agra", "Lucknow", "Kanpur", "Bhopal", "Gurgaon"];
  const distData = districts.map(d => ({
    name: d, count: req.filter(r => r.district === d).length,
    vol: sum(req.filter(r => r.district === d && r.status === "Approved"), "qty"),
    ncr: avg(req.filter(r => r.district === d), "ncrPmt"),
    freight: avg(req.filter(r => r.district === d), "primaryFreight"),
  })).sort((a, b) => b.vol - a.vol);

  const zoneData = ZONES.map(z => ({
    name: z, count: req.filter(r => r.zone === z).length,
    vol: sum(req.filter(r => r.zone === z && r.status === "Approved"), "qty"),
    ncr: avg(req.filter(r => r.zone === z), "ncrPmt"),
  }));

  const regionBar = REGIONS.map(r => ({
    name: r.replace("Uttar Pradesh", "UP").replace("Madhya Pradesh", "MP"),
    vol: sum(req.filter(x => x.region === r && x.status === "Approved"), "qty"),
    count: req.filter(x => x.region === r).length,
    ncr: avg(req.filter(x => x.region === r), "ncrPmt"),
  }));

  return (
    <div className="page">
      <div className="ph">
        <div>
          <div className="ph-ey">Geo Analytics</div>
          <div className="ph-title">Geographic Analytics Dashboard</div>
          <div className="ph-sub">Region · Zone · District · Freight impact analysis</div>
        </div>
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div className="card-title">Region-wise Volume (MT)</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={regionBar} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT />} />
                <Bar dataKey="vol" name="Volume MT" radius={[5,5,0,0]}>
                  {regionBar.map((_, i) => <Cell key={i} fill={CHART_COLORS[i]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
        <div className="card">
          <div className="card-hdr"><div className="card-title">Zone-wise NCR/MT</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={zoneData} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT fmt={v => `₹${v.toFixed(0)}`} />} />
                <Bar dataKey="ncr" name="Avg NCR/MT" radius={[5,5,0,0]}>
                  {zoneData.map((z, i) => <Cell key={i} fill={z.ncr >= 0 ? G.grn : G.red} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="card">
        <div className="card-hdr"><div><div className="card-title">District-wise Analysis</div><div className="card-sub">Volume, NCR, and freight impact per district</div></div></div>
        <div className="tw">
          <table>
            <thead><tr><th>District</th><th>Requests</th><th>Volume (MT)</th><th>Avg NCR/MT</th><th>Avg Freight</th><th>Volume Rank</th></tr></thead>
            <tbody>
              {distData.map((d, i) => (
                <tr key={d.name}>
                  <td style={{ fontWeight: 700 }}>{d.name}</td>
                  <td>{d.count}</td>
                  <td style={{ fontFamily: "'Sora',sans-serif", fontWeight: 700, color: G.blue }}>{d.vol.toLocaleString()} MT</td>
                  <td><span className={d.ncr < 0 ? "ncr-neg" : "ncr-pos"}>₹{d.ncr.toFixed(0)}</span></td>
                  <td style={{ fontWeight: 600 }}>₹{d.freight.toFixed(0)}/MT</td>
                  <td>
                    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                      <div style={{ width: 50, height: 6, background: G.border, borderRadius: 3, overflow: "hidden" }}>
                        <div style={{ height: "100%", width: `${pct(d.vol, distData[0].vol)}%`, background: CHART_COLORS[i % 8], borderRadius: 3 }} />
                      </div>
                      <span style={{ fontSize: 11, fontWeight: 700, color: G.muted }}>#{i + 1}</span>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ════════════════════════════════════════════════════════════════
   DASHBOARD 11 — AUDIT & COMPLIANCE
   ════════════════════════════════════════════════════════════════ */
function AuditDash({ req }) {
  const exceptions = req.filter(r => r.ncrPmt < -300 && r.status === "Approved");
  const excessDisc = req.filter(r => r.difference < -30 && r.status === "Approved");
  const blocked    = req.filter(r => r.blocked);

  const monthAudit = MONTHS.map(m => ({
    m,
    approved: req.filter(r => r.month === m && r.status === "Approved").length,
    exceptions: req.filter(r => r.month === m && r.ncrPmt < -300).length,
  }));

  return (
    <div className="page">
      <div className="ph">
        <div>
          <div className="ph-ey">Compliance</div>
          <div className="ph-title">Audit & Compliance Dashboard</div>
          <div className="ph-sub">Policy exceptions · Override tracking · Compliance score</div>
        </div>
        <div style={{ padding: "10px 18px", background: exceptions.length > 3 ? "#FEF2F2" : "#ECFDF5", borderRadius: 10, border: `1.5px solid ${exceptions.length > 3 ? "#FCA5A5" : "#A7F3D0"}` }}>
          <div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 24, color: exceptions.length > 3 ? G.red : G.grn }}>
            {Math.max(0, 100 - exceptions.length * 5)}%
          </div>
          <div style={{ fontSize: 11, fontWeight: 600, color: G.muted, textTransform: "uppercase" }}>Compliance Score</div>
        </div>
      </div>

      <div className="kpi-grid kpi-grid-4">
        <KPI icon="⚠️" val={exceptions.length}   label="NCR Policy Exceptions" trend="down" trendTxt="Approved <-300 NCR" color={G.red}  />
        <KPI icon="💸" val={excessDisc.length}    label="Excess Discount Cases"  trend="down" trendTxt="Discount > ₹30"     color={G.ora}  />
        <KPI icon="🔒" val={blocked.length}        label="Blocked Orders"         trend="neu" trendTxt="Admin blocked"       color={G.gold} />
        <KPI icon="📋" val={req.length}            label="Total Auditable Actions" trend="up" trendTxt="Complete trail"      color={G.blue} />
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div className="card-title">Exceptions vs Approvals Trend</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={190}>
              <LineChart data={monthAudit} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="m" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT />} />
                <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
                <Line type="monotone" dataKey="approved"   stroke={G.grn} strokeWidth={2} name="Approved" dot={{ r: 3 }} />
                <Line type="monotone" dataKey="exceptions" stroke={G.red} strokeWidth={2} name="Exceptions" dot={{ r: 3 }} strokeDasharray="5 3" />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>
        <div className="card">
          <div className="card-hdr"><div className="card-title">Policy Exception Breakdown</div></div>
          <div className="card-body">
            {[
              { label: "NCR below -300 threshold", count: exceptions.length,  color: G.red,   icon: "⚠️" },
              { label: "Excess discount (>₹30)",   count: excessDisc.length,  color: G.ora,   icon: "💸" },
              { label: "Blocked customer orders",  count: blocked.length,      color: G.gold,  icon: "🔒" },
              { label: "TAT exceeded (>7 days)",   count: req.filter(r => r.tat > 7).length, color: G.blue, icon: "⏱️" },
            ].map((e, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 0", borderBottom: `1px solid ${G.border}` }}>
                <span style={{ fontSize: 18 }}>{e.icon}</span>
                <div style={{ flex: 1, fontSize: 13, fontWeight: 600 }}>{e.label}</div>
                <span style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 18, color: e.color }}>{e.count}</span>
                <span className={`badge ${e.count > 3 ? "b-red" : e.count > 0 ? "b-ora" : "b-grn"}`}>{e.count > 3 ? "High" : e.count > 0 ? "Medium" : "Clear"}</span>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="card">
        <div className="card-hdr"><div><div className="card-title">🚨 Policy Exceptions — Approved with NCR &lt; -300</div><div className="card-sub">These require management review and justification</div></div></div>
        <div className="tw">
          <table>
            <thead><tr><th>Request ID</th><th>Customer</th><th>Grade</th><th>Region</th><th>NCR/MT</th><th>Approved By</th><th>Action</th></tr></thead>
            <tbody>
              {exceptions.slice(0, 8).map(r => (
                <tr key={r.id}>
                  <td style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: G.goldDk, fontSize: 12 }}>{r.id}</td>
                  <td style={{ fontWeight: 600 }}>{r.customer}</td>
                  <td><span className="badge b-muted">{r.grade}</span></td>
                  <td style={{ fontSize: 12 }}>{r.region}</td>
                  <td><span style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 14, color: G.red }}>₹{r.ncrPmt}</span></td>
                  <td style={{ fontSize: 12, color: G.muted }}>Level {rnd(2, 5)}</td>
                  <td><span className="badge b-red">🚨 Review</span></td>
                </tr>
              ))}
              {exceptions.length === 0 && <tr><td colSpan={7} style={{ textAlign: "center", padding: 30, color: G.grn, fontWeight: 700 }}>✅ No policy exceptions found</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ════════════════════════════════════════════════════════════════
   DASHBOARD 12 — MASTER DATA
   ════════════════════════════════════════════════════════════════ */
function MasterDash({ req }) {
  const custCount = 8;
  const activeC = 7; const inactiveC = 1;
  const custTrend = MONTHS.map((m, i) => ({ m, added: rnd(0, 2) }));

  const plantVol = PLANTS.map(p => ({
    name: p.replace(" Plant", ""), vol: sum(req.filter(r => r.plant === p && r.status === "Approved"), "qty"),
    count: req.filter(r => r.plant === p).length,
    util: rnd(55, 95),
  }));

  const regionCust = REGIONS.map(r => ({
    name: r.replace("Uttar Pradesh", "UP").replace("Madhya Pradesh", "MP"),
    count: rnd(1, 4),
  }));

  return (
    <div className="page">
      <div className="ph">
        <div>
          <div className="ph-ey">Master Data</div>
          <div className="ph-title">Master Data Dashboard</div>
          <div className="ph-sub">Customer · Plant · Supply source analytics</div>
        </div>
      </div>

      <div className="kpi-grid kpi-grid-4">
        <KPI icon="🏢" val={custCount}   label="Total Customers"   trend="up"  trendTxt="In database"      color={G.gold}   />
        <KPI icon="✅" val={activeC}     label="Active Customers"  trend="up"  trendTxt={`${pct(activeC,custCount)}% active`}   color={G.grn}    />
        <KPI icon="❌" val={inactiveC}   label="Inactive Customers" trend="neu" trendTxt="No recent orders"  color={G.muted}  />
        <KPI icon="🏗" val={PLANTS.length} label="Plants / Depots"  trend="neu" trendTxt="Supply sources"    color={G.blue}   />
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div className="card-title">Plant-wise Volume Dispatched</div></div>
          <div className="card-body">
            {plantVol.map((p, i) => (
              <div key={p.name} style={{ marginBottom: 18 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 5 }}>
                  <div>
                    <span style={{ fontWeight: 700, fontSize: 14 }}>{p.name}</span>
                    <span style={{ fontSize: 11, color: G.muted, marginLeft: 8, fontWeight: 600 }}>{p.count} requests</span>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <span style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, color: CHART_COLORS[i] }}>{p.vol.toLocaleString()} MT</span>
                    <span style={{ fontSize: 11, color: G.muted, marginLeft: 6 }}>Util: {p.util}%</span>
                  </div>
                </div>
                <div style={{ background: G.border, borderRadius: 6, height: 12, overflow: "hidden" }}>
                  <div style={{ height: "100%", width: `${pct(p.vol, Math.max(...plantVol.map(x => x.vol)))}%`, background: CHART_COLORS[i], borderRadius: 6, transition: "width .4s" }} />
                </div>
              </div>
            ))}
          </div>
        </div>

        <div className="card">
          <div className="card-hdr"><div className="card-title">Region-wise Customer Count</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={210}>
              <BarChart data={regionCust} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT />} />
                <Bar dataKey="count" name="Customers" radius={[5,5,0,0]}>
                  {regionCust.map((_, i) => <Cell key={i} fill={CHART_COLORS[i]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="card">
        <div className="card-hdr"><div className="card-title">Supply Source Trend (Plant-wise Requests)</div></div>
        <div className="card-body">
          <ResponsiveContainer width="100%" height={180}>
            <BarChart data={MONTHS.map(m => ({
              m,
              aligarh: req.filter(r => r.month === m && r.plant === "Aligarh Plant").length,
              morak: req.filter(r => r.month === m && r.plant === "Morak Plant").length,
              depot: req.filter(r => r.month === m && r.plant === "Depot").length,
            }))} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
              <XAxis dataKey="m" tick={{ fontSize: 11 }} />
              <YAxis tick={{ fontSize: 11 }} />
              <Tooltip content={<TT />} />
              <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
              <Bar dataKey="aligarh" stackId="a" fill={G.gold}   name="Aligarh" />
              <Bar dataKey="morak"   stackId="a" fill={G.blue}   name="Morak" />
              <Bar dataKey="depot"   stackId="a" fill={G.purple} name="Depot" radius={[3,3,0,0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

/* ════════════════════════════════════════════════════════════════
   DASHBOARD 13 — FORECAST & PLANNING
   ════════════════════════════════════════════════════════════════ */
function ForecastDash({ req }) {
  const approved = req.filter(r => r.status === "Approved");
  const pipeline = req.filter(r => r.status === "Pending");

  const forecastMonths = ["Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const forecastData = forecastMonths.map((m, i) => ({
    m, projected: rnd(800, 2000), pipeline: rnd(300, 800),
  }));

  const regionForecast = REGIONS.map(r => ({
    name: r.replace("Uttar Pradesh", "UP").replace("Madhya Pradesh", "MP"),
    pending: sum(req.filter(x => x.region === r && x.status === "Pending"), "qty"),
    approved: sum(req.filter(x => x.region === r && x.status === "Approved"), "qty"),
  }));

  const plantPlan = PLANTS.map(p => ({
    name: p.replace(" Plant", ""),
    pending: sum(req.filter(r => r.plant === p && r.status === "Pending"), "qty"),
    projected: rnd(1500, 4000),
  }));

  return (
    <div className="page">
      <div className="ph">
        <div>
          <div className="ph-ey">Planning</div>
          <div className="ph-title">Sales Forecast & Planning</div>
          <div className="ph-sub">Pipeline · Demand forecast · Plant dispatch planning</div>
        </div>
      </div>

      <div className="kpi-grid kpi-grid-4">
        <KPI icon="🔮" val={`${(sum(pipeline,"qty")/1000).toFixed(1)}K`} label="Pipeline Volume (MT)"  trend="up" trendTxt="Pending approvals"  color={G.gold}  />
        <KPI icon="📊" val={pipeline.length}   label="Active Pipeline"      trend="up" trendTxt="Requests pending"    color={G.blue}  />
        <KPI icon="🏭" val={`${(sum(approved.slice(-10),"qty")/1000).toFixed(1)}K`} label="Approvals This Month" trend="up" trendTxt="MT approved"  color={G.grn}   />
        <KPI icon="🎯" val={`${avg(approved,"qty").toFixed(0)} MT`} label="Avg Deal Size" trend="neu" trendTxt="Per approval"        color={G.purple} />
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-hdr"><div><div className="card-title">6-Month Volume Forecast</div><div className="card-sub">Based on pipeline and historical trends</div></div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={200}>
              <AreaChart data={forecastData} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <defs>
                  <linearGradient id="projGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%"  stopColor={G.gold} stopOpacity={0.3} />
                    <stop offset="95%" stopColor={G.gold} stopOpacity={0.0} />
                  </linearGradient>
                  <linearGradient id="pipeGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%"  stopColor={G.blue} stopOpacity={0.2} />
                    <stop offset="95%" stopColor={G.blue} stopOpacity={0.0} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="m" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT fmt={v => `${v} MT`} />} />
                <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
                <Area type="monotone" dataKey="projected" stroke={G.gold} fill="url(#projGrad)" strokeWidth={2} name="Projected (MT)" dot={{ r: 3 }} />
                <Area type="monotone" dataKey="pipeline"  stroke={G.blue} fill="url(#pipeGrad)" strokeWidth={2} name="Pipeline (MT)"  dot={{ r: 3 }} />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="card">
          <div className="card-hdr"><div className="card-title">Region Demand Forecast</div></div>
          <div className="card-body">
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={regionForecast} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={G.border} />
                <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip content={<TT />} />
                <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} />
                <Bar dataKey="approved" fill={G.grn}  name="Approved (MT)" radius={[0,0,0,0]} />
                <Bar dataKey="pending"  fill={G.gold} name="Pipeline (MT)"  radius={[3,3,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="card">
        <div className="card-hdr"><div><div className="card-title">Plant Dispatch Planning</div><div className="card-sub">Pending volume by supply source — dispatch priority</div></div></div>
        <div className="card-body">
          {plantPlan.map((p, i) => (
            <div key={p.name} style={{ marginBottom: 20 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                <div>
                  <span style={{ fontWeight: 700, fontSize: 15 }}>{p.name}</span>
                </div>
                <div style={{ display: "flex", gap: 20 }}>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 16, color: G.ora }}>{p.pending.toLocaleString()} MT</div>
                    <div style={{ fontSize: 10.5, color: G.muted, fontWeight: 600 }}>PENDING</div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontFamily: "'Sora',sans-serif", fontWeight: 800, fontSize: 16, color: G.gold }}>{p.projected.toLocaleString()} MT</div>
                    <div style={{ fontSize: 10.5, color: G.muted, fontWeight: 600 }}>PROJECTED</div>
                  </div>
                </div>
              </div>
              <div style={{ background: G.border, borderRadius: 6, height: 14, overflow: "hidden" }}>
                <div style={{ height: "100%", width: `${pct(p.pending, p.projected)}%`, background: CHART_COLORS[i], borderRadius: 6, transition: "width .4s" }} />
              </div>
              <div style={{ fontSize: 11, color: G.muted, marginTop: 4 }}>Utilization: {pct(p.pending, p.projected)}% of projected capacity</div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}



// ══ ANALYTICS DASHBOARD ROUTER ══
function AnalyticsDashboardRouter({ dashKey, requests }) {
  // ── Global filter state ────────────────────────────────────
  const [fZone,     setFZone]     = useState("");
  const [fRegion,   setFRegion]   = useState("");
  const [fDistrict, setFDistrict] = useState("");
  const [fGrade,    setFGrade]    = useState("");
  const [fPlant,    setFPlant]    = useState("");
  const [fPayTerm,  setFPayTerm]  = useState("");
  const [fStatus,   setFStatus]   = useState("");
  const [showFilters, setShowFilters] = useState(true);

  // ── Convert ntpas_v5 format → analytics format ────────────
  const allReq = useMemo(() => {
    if (!requests || !requests.length) return [];
    return requests.map(r => {
      const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
      const monthIdx = r.date ? new Date(r.date).getMonth() : 0;
      return {
        id: r.id,
        month: months[Math.min(monthIdx, 5)],
        monthIdx: Math.min(monthIdx, 5),
        grade: r.grade || "OPC",
        plant: r.supplyFrom || "Aligarh Plant",
        region: r.region || "Rajasthan",
        zone: r.zone || "Zone A",
        district: r.destination || r.district || "Jaipur",
        pay: r.paymentTerms || "Advance",
        orderPrice: Number(r.orderPrice || 0),
        tradePrice: Number(r.tradePrice || 0),
        difference: Number(r.difference || 0),
        netOfGST: Number(r.netOfGST || 0),
        primaryFreight: Number(r.primaryFreight || 0),
        secondaryFreight: Number(r.secondaryFreight || 0),
        cop: Number(r.costOfProduction || 210),
        packing: Number(r.packing || 25),
        opCommission: Number(r.opCommission || 5),
        totalExpenses: Number(r.totalExpenses || 0),
        ncrPmt: Number(r.ncrPmt || 0),
        qty: Number(r.qty || 0),
        status: r.status || "Pending",
        currentLevel: r.currentLevel || "Sales Officer",
        tat: r.tat || (r.history ? r.history.length : 1),
        invoicedQty: Number(r.invoicedQty || 0),
        customer: r.customerName || "Customer",
        soId: r.createdBy || "SO001",
        soName: r.createdByName || "Sales Officer",
        blocked: !!r.blocked,
        history: r.history || [],
      };
    });
  }, [requests]);

  // ── Derive unique dropdown values from actual data ─────────
  const zones     = useMemo(() => [...new Set(allReq.map(r => r.zone).filter(Boolean))].sort(), [allReq]);
  const regions   = useMemo(() => [...new Set(allReq.map(r => r.region).filter(Boolean))].sort(), [allReq]);
  const districts = useMemo(() => [...new Set(allReq.map(r => r.district).filter(Boolean))].sort(), [allReq]);
  const grades    = useMemo(() => [...new Set(allReq.map(r => r.grade).filter(Boolean))].sort(), [allReq]);
  const plants    = useMemo(() => [...new Set(allReq.map(r => r.plant).filter(Boolean))].sort(), [allReq]);
  const payTerms  = useMemo(() => [...new Set(allReq.map(r => r.pay).filter(Boolean))].sort(), [allReq]);

  // ── Apply all filters ──────────────────────────────────────
  const req = useMemo(() => allReq.filter(r =>
    (!fZone     || r.zone     === fZone)     &&
    (!fRegion   || r.region   === fRegion)   &&
    (!fDistrict || r.district === fDistrict) &&
    (!fGrade    || r.grade    === fGrade)    &&
    (!fPlant    || r.plant    === fPlant)    &&
    (!fPayTerm  || r.pay      === fPayTerm)  &&
    (!fStatus   || r.status   === fStatus)
  ), [allReq, fZone, fRegion, fDistrict, fGrade, fPlant, fPayTerm, fStatus]);

  const activeFilters = [fZone,fRegion,fDistrict,fGrade,fPlant,fPayTerm,fStatus].filter(Boolean).length;

  const clearAll = () => { setFZone(""); setFRegion(""); setFDistrict(""); setFGrade(""); setFPlant(""); setFPayTerm(""); setFStatus(""); };

  // ── Find current dashboard title ────────────────────────────
  const dashInfo = ANALYTICS_DASHBOARDS.find(d => d.key === dashKey) || { label: dashKey, icon: "📊" };

  // ── Dashboard map ──────────────────────────────────────────
  const DASH_MAP = {
    exec:     <ExecDash     req={req} />,
    workflow: <WorkflowDash req={req} />,
    ncr:      <NCRDash      req={req} />,
    perf:     <PerfDash     req={req} />,
    sales:    <SalesDash    req={req} />,
    customer: <CustomerDash req={req} />,
    geo:      <GeoDash      req={req} />,
    freight:  <FreightDash  req={req} />,
    aging:    <AgingDash    req={req} />,
    audit:    <AuditDash    req={req} />,
    trade:    <TradeDash    req={req} />,
    master:   <MasterDash   req={req} />,
    forecast: <ForecastDash req={req} />,
  };

  const dash = DASH_MAP[dashKey];
  if (!dash) return (
    <div style={{padding:40,textAlign:"center",color:"#999"}}>
      <div style={{fontSize:40,marginBottom:12}}>📊</div>
      <div style={{fontWeight:700,fontSize:18}}>Dashboard not found: {dashKey}</div>
    </div>
  );

  // ── Filter bar styles ──────────────────────────────────────
  const selStyle = {
    padding:"7px 11px", border:"1.5px solid #E7E5E4", borderRadius:8,
    fontSize:12.5, fontFamily:"inherit", background:"#fff", color:"#1C1917",
    fontWeight:600, cursor:"pointer", outline:"none", minWidth:130,
    transition:"border-color .15s",
  };
  const activeSel = { ...selStyle, borderColor:"#C9A84C", background:"#FFF9EE" };

  return (
    <div style={{background:"#FAFAF9",minHeight:"100%"}}>
      <style>{css}</style>

      {/* ── Global Filter Bar ────────────────────────────── */}
      <div style={{
        background:"#fff", borderBottom:"2px solid #E7E5E4",
        padding:"0 24px", position:"sticky", top:0, zIndex:200,
        boxShadow:"0 2px 8px rgba(0,0,0,.06)",
      }}>
        {/* Header row */}
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 0 8px"}}>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <span style={{fontSize:20}}>{dashInfo.icon}</span>
            <div>
              <div style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:16,color:"#1C1917"}}>{dashInfo.label}</div>
              <div style={{fontSize:11,color:"#78716C",fontWeight:600}}>
                Showing <strong style={{color: activeFilters>0?"#C9A84C":"#059669"}}>{req.length}</strong> of <strong>{allReq.length}</strong> records
                {activeFilters > 0 && <span style={{color:"#C9A84C"}}> · {activeFilters} filter{activeFilters>1?"s":""} active</span>}
              </div>
            </div>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            {activeFilters > 0 && (
              <button onClick={clearAll} style={{
                fontSize:11.5,fontWeight:700,color:"#DC2626",background:"#FEF2F2",
                border:"1.5px solid #FECACA",borderRadius:7,padding:"5px 12px",cursor:"pointer"
              }}>✕ Clear All Filters</button>
            )}
            <button onClick={() => setShowFilters(f=>!f)} style={{
              fontSize:11.5,fontWeight:700,color:"#78716C",background:"#F5F5F4",
              border:"1.5px solid #E7E5E4",borderRadius:7,padding:"5px 12px",cursor:"pointer",
              display:"flex",alignItems:"center",gap:5
            }}>
              <span>⚙</span> {showFilters?"Hide":"Show"} Filters
              {activeFilters > 0 && <span style={{background:"#C9A84C",color:"#fff",borderRadius:10,padding:"1px 6px",fontSize:10}}>{activeFilters}</span>}
            </button>
          </div>
        </div>

        {/* Filter selects row */}
        {showFilters && (
          <div style={{display:"flex",gap:10,flexWrap:"wrap",paddingBottom:12,alignItems:"center"}}>
            {/* Zone */}
            <div style={{display:"flex",flexDirection:"column",gap:3}}>
              <label style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:.6,color:"#78716C"}}>Zone</label>
              <select style={fZone ? activeSel : selStyle} value={fZone} onChange={e=>setFZone(e.target.value)}>
                <option value="">All Zones</option>
                {zones.map(z=><option key={z}>{z}</option>)}
              </select>
            </div>
            {/* Region */}
            <div style={{display:"flex",flexDirection:"column",gap:3}}>
              <label style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:.6,color:"#78716C"}}>Region</label>
              <select style={fRegion ? activeSel : selStyle} value={fRegion} onChange={e=>setFRegion(e.target.value)}>
                <option value="">All Regions</option>
                {regions.map(r=><option key={r}>{r}</option>)}
              </select>
            </div>
            {/* District */}
            <div style={{display:"flex",flexDirection:"column",gap:3}}>
              <label style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:.6,color:"#78716C"}}>District</label>
              <select style={fDistrict ? activeSel : selStyle} value={fDistrict} onChange={e=>setFDistrict(e.target.value)}>
                <option value="">All Districts</option>
                {districts.map(d=><option key={d}>{d}</option>)}
              </select>
            </div>
            {/* Material Grade */}
            <div style={{display:"flex",flexDirection:"column",gap:3}}>
              <label style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:.6,color:"#78716C"}}>Material Grade</label>
              <select style={fGrade ? activeSel : selStyle} value={fGrade} onChange={e=>setFGrade(e.target.value)}>
                <option value="">All Grades</option>
                {grades.map(g=><option key={g}>{g}</option>)}
              </select>
            </div>
            {/* Plant / Supply Source */}
            <div style={{display:"flex",flexDirection:"column",gap:3}}>
              <label style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:.6,color:"#78716C"}}>Supply Plant</label>
              <select style={fPlant ? activeSel : selStyle} value={fPlant} onChange={e=>setFPlant(e.target.value)}>
                <option value="">All Plants</option>
                {plants.map(p=><option key={p}>{p}</option>)}
              </select>
            </div>
            {/* Payment Terms */}
            <div style={{display:"flex",flexDirection:"column",gap:3}}>
              <label style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:.6,color:"#78716C"}}>Payment Terms</label>
              <select style={fPayTerm ? activeSel : selStyle} value={fPayTerm} onChange={e=>setFPayTerm(e.target.value)}>
                <option value="">All Terms</option>
                {payTerms.map(p=><option key={p}>{p}</option>)}
              </select>
            </div>
            {/* Status */}
            <div style={{display:"flex",flexDirection:"column",gap:3}}>
              <label style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:.6,color:"#78716C"}}>Status</label>
              <select style={fStatus ? activeSel : selStyle} value={fStatus} onChange={e=>setFStatus(e.target.value)}>
                <option value="">All Statuses</option>
                <option>Approved</option>
                <option>Pending</option>
                <option>Rejected</option>
                <option>Blocked</option>
              </select>
            </div>

            {/* Active filter chips */}
            {activeFilters > 0 && (
              <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"flex-end",marginLeft:4}}>
                {[
                  {label:"Zone",val:fZone,clear:()=>setFZone("")},
                  {label:"Region",val:fRegion,clear:()=>setFRegion("")},
                  {label:"District",val:fDistrict,clear:()=>setFDistrict("")},
                  {label:"Grade",val:fGrade,clear:()=>setFGrade("")},
                  {label:"Plant",val:fPlant,clear:()=>setFPlant("")},
                  {label:"Payment",val:fPayTerm,clear:()=>setFPayTerm("")},
                  {label:"Status",val:fStatus,clear:()=>setFStatus("")},
                ].filter(f=>f.val).map(f=>(
                  <div key={f.label} style={{
                    display:"flex",alignItems:"center",gap:5,
                    background:"#FFF9EE",border:"1.5px solid #F0D080",
                    borderRadius:20,padding:"3px 10px",fontSize:11.5,fontWeight:700,color:"#A07838"
                  }}>
                    <span style={{color:"#78716C",fontWeight:600}}>{f.label}:</span> {f.val}
                    <button onClick={f.clear} style={{background:"none",border:"none",cursor:"pointer",color:"#A07838",padding:0,fontSize:12,fontWeight:800,lineHeight:1}}>×</button>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* Zero results warning */}
        {req.length === 0 && activeFilters > 0 && (
          <div style={{padding:"8px 0 12px",color:"#DC2626",fontWeight:700,fontSize:13,display:"flex",gap:8,alignItems:"center"}}>
            <span>⚠️</span> No records match the selected filters.
            <button onClick={clearAll} style={{fontSize:12,fontWeight:700,color:"#DC2626",background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>Clear filters</button>
          </div>
        )}
      </div>

      {/* ── Dashboard content ──────────────────────────── */}
      {dash}
    </div>
  );
}


// ═══ APP ROOT ══════════════════════════════════════════════
export default function App() {
  const [user, setUser] = useState(null);
  const [page, setPage] = useState("dashboard");
  const [activeTab, setActiveTab] = useState("all");
  const [requests, setRequests] = useState(INITIAL_REQUESTS);
  const [customers, setCustomers] = useState(INITIAL_CUSTOMERS);
  const [priceMaster, setPriceMaster]       = useState(INITIAL_PRICE_MASTER);
  const [freightMaster, setFreightMaster]   = useState(INITIAL_FREIGHT_MASTER);
  const [locationMaster, setLocationMaster] = useState(INITIAL_LOCATION_MASTER);
  const [copMaster,     setCopMaster]     = useState(INITIAL_COP_MASTER);
  const [packingMaster, setPackingMaster] = useState(INITIAL_PACKING_MASTER);
  const [plantMaster,   setPlantMaster]   = useState(INITIAL_PLANT_MASTER);
  const [tpcAgents, setTpcAgents] = useState(INITIAL_TPC_AGENTS);
  const [modeMaster, setModeMaster] = useState(INITIAL_MODE_MASTER);
  const [unitSourceMaster, setUnitSourceMaster] = useState(INITIAL_UNIT_SOURCE_MASTER);
  const [sourceMaster, setSourceMaster] = useState(INITIAL_SOURCE_MASTER);
  const [storageLocationMaster, setStorageLocationMaster] = useState(INITIAL_STORAGE_LOCATION_MASTER);
  const [salesEntries, setSalesEntries] = useState(INITIAL_SALES_ENTRIES);
  const [users, setUsers] = useState(INITIAL_USERS);
  const [opCommissionMaster, setOpCommissionMaster] = useState(INITIAL_OP_COMMISSION_MASTER);

  useEffect(() => { if (user) { const u = users.find(u => u.id === user.id); if (u && (u.name !== user.name || u.role !== user.role)) setUser(u); } }, [users]);

  const handleAction = (reqId, actionType, remark, modPrice, modQty, modValidFrom, modValidTo, modPaymentTerms) => {
    setRequests(prev => prev.map(r => {
      if (r.id !== reqId) return r;
      const now = new Date().toLocaleString();

      // Block / Unblock — admin only
      if (actionType === "block") {
        return { ...r, blocked: true, blockReason: remark, blockedAt: now, blockedBy: user.id,
          history: [...r.history, { level: user.role, action: "Blocked", by: user.id, time: now, remark }] };
      }
      if (actionType === "unblock") {
        return { ...r, blocked: false, blockReason: "", blockedAt: "", blockedBy: "",
          history: [...r.history, { level: user.role, action: "Unblocked", by: user.id, time: now, remark }] };
      }

      // Resubmit — Sales Officer edits and sends back up to ASM
      if (actionType === "resubmit") {
        const newH = [...r.history, { level: user.role, action: "Resubmitted", by: user.id, time: now, remark }];
        return {
          ...r,
          status: "Pending",
          currentLevel: "Area Sales Manager",
          history: newH,
          orderPrice:   modPrice        || r.orderPrice,
          qty:          modQty          || r.qty,
          validityFrom: modValidFrom    || r.validityFrom,
          validityTo:   modValidTo      || r.validityTo,
          paymentTerms: modPaymentTerms || r.paymentTerms,
          // recalculate difference and netOfGST
          difference: ((modPrice||r.orderPrice) - r.tradePrice).toFixed(2),
          netOfGST:   ((modPrice||r.orderPrice) / 1.18).toFixed(2),
        };
      }

      const actions = { approve: "Approved", reject: "Rejected", modify: "Modified", escalate: "Escalated", sendback: "Sent Back" };
      const newH = [...r.history, { level: user.role, action: actions[actionType] || actionType, by: user.id, time: now, remark }];
      let ns = r.status, nl = r.currentLevel;

      if (actionType === "approve") {
        const nx = nextLevel(r.currentLevel);
        if (!nx || r.currentLevel === "Sales & Accounts") { ns = "Approved"; nl = "Admin"; }
        else { ns = "Pending"; nl = nx; }
      } else if (actionType === "reject") {
        ns = "Rejected"; nl = "Regional Head";
      } else if (actionType === "sendback") {
        ns = "Pending"; nl = prevLevel(r.currentLevel);
      }

      return {
        ...r,
        status: ns,
        currentLevel: nl,
        history: newH,
        orderPrice: modPrice || r.orderPrice,
        qty: modQty || r.qty,
        validityFrom: (actionType === "modify" && modValidFrom) ? modValidFrom : r.validityFrom,
        validityTo:   (actionType === "modify" && modValidTo)   ? modValidTo   : r.validityTo,
      };
    }));
  };
  const handleNewRequest = req => { setRequests(p => [req, ...p]); setPage("requests"); setActiveTab("pending"); };

  // Location-based access filter: if user has allowedSalesOffices, restrict visible requests
  if (!user) return <LoginPage users={users} onLogin={u => { setUser(u); setPage("dashboard"); }} />;

  const currentUser = users.find(u => u.id === user.id) || user;
  const isFullAccess = user.role === "Admin" || user.role === "Zonal Head" || user.role === "Sales & Accounts";

  // Build allowed name sets from each location level
  const resolve = (field, lvl) => (currentUser[field]||[]).map(id=>(locationMaster[lvl]||[]).find(x=>x.id===id)?.name).filter(Boolean);
  const allowedRegionNames   = resolve("allowedRegions",    "regions");
  const allowedClusterNames  = resolve("allowedClusters",   "clusters");
  const allowedSONames       = resolve("allowedSalesOffices","salesOffices");
  const allowedDistNames     = resolve("allowedDistricts",  "districts");
  const allowedTehsilNames   = resolve("allowedTehsils",    "tehsils");
  const allowedCityNames     = resolve("allowedCities",     "cities");

  const hasAnyRestriction = !isFullAccess && (
    allowedRegionNames.length>0||allowedClusterNames.length>0||allowedSONames.length>0||
    allowedDistNames.length>0||allowedTehsilNames.length>0||allowedCityNames.length>0
  );

  const visibleRequests = (!hasAnyRestriction)
    ? requests
    : requests.filter(r => {
        const dest   = String(r.destination||"");
        const reg    = String(r.region||"");
        const zone   = String(r.zone||"");
        // Pass if ANY allowed level matches the request
        const passRegion  = allowedRegionNames.length===0  || allowedRegionNames.some(n=>reg.includes(n)||dest.includes(n));
        const passCluster = allowedClusterNames.length===0 || allowedClusterNames.some(n=>zone.includes(n)||dest.includes(n));
        const passSO      = allowedSONames.length===0      || allowedSONames.some(n=>dest.includes(n)||reg.includes(n));
        const passDist    = allowedDistNames.length===0    || allowedDistNames.some(n=>dest.includes(n));
        const passTehsil  = allowedTehsilNames.length===0  || allowedTehsilNames.some(n=>dest.includes(n));
        const passCity    = allowedCityNames.length===0    || allowedCityNames.some(n=>dest.includes(n));
        // A request is visible if it passes EVERY configured restriction level
        return passRegion && passCluster && passSO && passDist && passTehsil && passCity;
      });

  const pendingCount = requests.filter(r => r.currentLevel === user.role && r.status === "Pending" && (user.role === "Admin" || user.role === "Sales & Accounts" || user.role === "Zonal Head" || (users.find(u => u.id === user.id)?.assignedCustomers || []).includes(r.customerCode))).length;

  // Pages the current user is allowed to access (empty = all)
  const allowedPages = currentUser.allowedPages||[];
  const canAccessPage = (key) => user.role==="Admin" || allowedPages.length===0 || allowedPages.includes(key);

  const NAV = [
    { key: "dashboard", icon: "📊", label: "Dashboard", group: "main" },
    ...(canAccessPage("new-request") && (user.role === "Sales Officer" || user.role === "Admin") ? [{ key: "new-request", icon: "➕", label: "New Request", group: "main" }] : []),
    ...(canAccessPage("requests")       ? [{ key: "requests",        icon: "📋", label: "Price Requests", badge: pendingCount||null, group: "main" }] : []),
    ...(canAccessPage("sale-updation")  ? [{ key: "sale-updation",   icon: "📝", label: "Sale Updation",  group: "main" }] : []),
    ...(canAccessPage("reports")        ? [{ key: "reports",         icon: "📈", label: "Price Report",   group: "analytics" }] : []),
    ...(canAccessPage("sales-report")   ? [{ key: "sales-report",    icon: "📊", label: "Sales Report",   group: "analytics" }] : []),
    ...(canAccessPage("ncr-calculator") ? [{ key: "ncr-calculator",  icon: "🧮", label: "NCR Calculator", group: "analytics" }] : []),
    ...(canAccessPage("audit")          ? [{ key: "audit",           icon: "📜", label: "Audit Log",      group: "analytics" }] : []),
    ...(user.role === "Admin" ? [{ key: "customer-master", icon: "👥", label: "Customer Master", group: "admin" }, { key: "admin-masters", icon: "⚙️", label: "Admin Masters", group: "admin" }] : []),
    ...ANALYTICS_DASHBOARDS.filter(d => user.role === "Admin" || (currentUser.analyticsRights || []).includes(d.key))
      .map(d => ({ key: `adash-${d.key}`, icon: d.icon, label: d.label, group: "dashboards" })),
  ];
  const groups = { main: "Navigation", analytics: "Analytics", admin: "Administration", dashboards: "Analytics Dashboards" };

  return (
    <>
      <style>{CSS}</style>
      <header className="hdr">
        <div className="hdr-logo">MC</div>
        <div className="hdr-brand"><div style={{overflow:"hidden",borderRadius:8,background:"#fff",padding:"4px 8px",marginBottom:3,maxWidth:160}}><img src="data:image/png;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCANqBO4DASIAAhEBAxEB/8QAHQABAAIDAQEBAQAAAAAAAAAAAAYIBAUHAwIBCf/EAGoQAAEDAgIBCg4MCAsFBgUEAwABAgMEBQYRBxITFiExQVFWYZMIFBc1VFVxdJGSlLLR0hgiMkJjcnOBsbPB0xUzNlJioaLhIzQ3RlNldYKVtMMng4SFwiRDZKPj8AkmRWalJUR2pEdXxP/EABwBAQEAAgMBAQAAAAAAAAAAAAABAwUCBAcGCP/EAEIRAQABAgIDDAgFBAICAwEBAAABAgMFEQQWURITFBUxM1JTYXGR0gYhMjRBgbHRB0KCocIiNcHwcrIjohdEkuHx/9oADAMBAAIRAxEAPwDl9jpKVbLQqtNCqrTRqqrGn5qGZ0nSdiwc2h5WLrJQd7R+ahmHzVUznL9F6Nao3mj1RyR9Hh0nSdiwc2g6TpOxYObQ9wcc5Zt6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0HSdJ2LBzaHuBnJvVGyHh0nSdiwc2g6TpOxYObQ9wM5N6o2Q8Ok6TsWDm0I3pEpqdlkhVkETV6ZamaMRPeuJWRnSP1jh75b5rjLZmd8hrMZt0RoN3KI5G6sXWSg72j81DMMOxdZKDvaPzUMwx1e1LYaNzNHdH0AAcWcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjOkfrHD3y3zXEmIzpH6xw98t81xlsc5DV417hd7m6sXWSg72j81DMMOxdZKDvaPzUMw4Ve1Lu6NzNHdH0AAcWcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjOkfrHD3y3zXEmIzpH6xw98t81xlsc5DV417hd7m6sXWSg72j81DMMOxdZKDvaPzUMw4Ve1Lu6NzNHdH0AAcWcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjOkfrHD3y3zXEmIzpH6xw98t81xlsc5DV417hd7m6sXWSg72j81DMMOxdZKDvaPzUMw4Ve1Lu6NzNHdH0AAcWcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjOkfrHD3y3zXEmIzpH6xw98t81xlsc5DV417hd7m6sXWSg72j81DMMOxdZKDvaPzUMw4Ve1Lu6NzNHdH0AAcWcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjOkfrHD3y3zXEmIzpH6xw98t81xlsc5DV417hd7m6sXWSg72j81DMMOxdZKDvaPzUMw4Ve1Lu6NzNHdH0AAcWcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjOkfrHD3y3zXEmIzpH6xw98t81xlsc5DV417hd7m6sXWSg72j81DMMOxdZKDvaPzUMw4Ve1Lu6NzNHdH0AAcWcAAAAAAAAAAAAAAAAAAAAsV0N2jOidbIcZX6lZUTTKq2+CVubY2ouWuqi7rlVNrgTb30yy2bNV2rcw1eL4tZwvR5v3fX8IjbOxyGw6NsdXymbU23DVa+B6Zskl1MLXJwosipmnKhtOovpM4sr5bT/eFxwbSMOt/GZecXPT7T5q/ot0xHbnP+Y+im/UX0mcWV8tp/vB1F9JnFlfLaf7wuQC8X29ssWvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMpv1F9JnFlfLaf7wdRfSZxZXy2n+8LkAcX29smvmI9Cjwq8ym/UX0mcWV8tp/vB1F9JnFlfLaf7wuQBxfb2ya+Yj0KPCrzKb9RfSZxZXy2n+8HUX0mcWV8tp/vC5AHF9vbJr5iPQo8KvMps7QzpMa1XLhh2ScFZTr/qESxDh++4dqUp75aau3yO9zr0StR/xV3HfMX2MC/2a2X61TWu70UVZSTJk+ORM/nRd1FTeVNtDjVh9GX9Muxo3p7pUVxv9umaezOJ/eZUFRx+opKNMGDJsB4ymtWrfLRSt1+imduviVVTJf0kVFRe5nvkSY/M1tdqaJyl6HoeJ2tKoiuic4l7A/EU/TC2kTmAAKAAAAAAAAAAAAAAAAEZ0j9Y4e+W+a4kxGdI/WOHvlvmuMtjnIavGvcLvc3Vi6yUHe0fmoZhh2LrJQd7R+ahmHCr2pd3RuZo7o+gADizgAAAAAAAAAAAAAAAAAA/FL74apoqLDlso4Wo2KCkijYibyNYiIUIcX+tPWqk+QZ5qG0w6PXU82/EGqdxYjtq/wygAbR5mAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACvfRo08X4Hw5XalNdZUTRIv6LmtXLwtQrbC/Mst0aq5YWw/wB/SfVlYoHHQ0qjOc32vo5plVFEUZtjGuZ6oY8Knu01NcZS9U0S5u6Il+gA4O4AAAAAAAAAAAAAAAAEZ0j9Y4e+W+a4kxGdI/WOHvlvmuMtjnIavGvcLvc3Vi6yUHe0fmoZhh2LrJQd7R+ahmHCr2pd3RuZo7o+gADizgAAAAAAAAAAAAAAAAAA+XF/7T1qpPkGeahQBxf+09aqT5BnmobTDvzfJ5p+IPJY/V/FlAA2jzUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFfujW/JXD/f0n1ZWCnUs/0a/wCSuH+/pPqysFOdTSH0uB8sNhDuGQ0x4NwyGmnucr17D+bh+gAxNkAAAAAAAAAAAAAAAAEZ0j9Y4e+W+a4kxGdI/WOHvlvmuMtjnIavGvcLvc3Vi6yUHe0fmoZhh2LrJQd7R+ahmHCr2pd3RuZo7o+gADizgAAAAAAAAAAAAAAAAAA+XF/7T1qpPkGeahQBxf8AtPWqk+QZ5qG0w783yeafiDyWP1fxZQANo81AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABX7o1/yVw/39J9WVgpyz/Rr/krh/v6T6srBTnU0h9JgfLDYQbhkNMeDcMhpp7nK9fw/m4foAMTZAAAAAAAAAAAAAAAABGdI/WOHvlvmuJMRnSP1jh75b5rjLY5yGrxr3C73N1YuslB3tH5qGYYdi6yUHe0fmoZhwq9qXd0bmaO6PoAA4s4AAAAAAAAAAAAAAAAAAPlxf+09aqT5BnmoUAcX/tPWqk+QZ5qG0w783yeafiDyWP1fxZQANo81AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGuuV8stsVUuV4t9Flu9MVLI8vGVDRVWkzANMqpJiy1uy/o5tc83M4zXTTyy7NrQ9IvRnbtzPdEylwIK7S9o4RclxRT/NDKv/SekOlfR5KuTcU0afHa9v0tQ479b6UeLPOE6fHrmxX/APmfsmwI9RY3wbWqjaXFVlkcu43p2NHeBVzN7TzwVEaS080czF3HMcjkX50OcVRPJLqXLF21zlMx3xk9AAViAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABX7o1/yVw/39J9WVgpyz/Rr/krh/v6T6srBTnU0h9JgfLDYQbhkNMeDcMhpp7nK9fw/m4foAMTZAAAAAAAAAAAAAAAABGdI/WOHvlvmuJMRnSP1jh75b5rjLY5yGrxr3C73N1YuslB3tH5qGYYdi6yUHe0fmoZhwq9qXd0bmaO6PoAA4s4AAAAAAAAAAAAAAAAAAPlxf8AtPWqk+QZ5qFAHF/7T1qpPkGeahtMO/N8nmn4g8lj9X8WUADaPNQAAAAAAAAAAAAAANRiHEthw8sCXq6U9Cs+q1rXXZavU5Z5dzNPCBtwRPqkYE40W7x/3DqkYE40W7x/3ASwET6pGBONFu8f9w6pGBONFu8f9wEsBE+qRgTjRbvH/cfUWkXBEsrIo8TW9z3uRrWo/bVV3E3AJUAazEF/s1ggjnvNxgoYpXahjpVyRzss8gNmCJ9UjAnGi3eP+4dUjAnGi3eP+4CWAifVIwJxot3j/uHVIwJxot3j/uAlgIn1SMCcaLd4/wC4dUjAnGi3eP8AuAlgPmN7ZGNexUc1yIqKm+hi3i52+z2+S4XSrjpKWNUR8si5NbmqIn61QDMBE+qRgTjRbvH/AHDqkYE40W7x/wBwEsBE+qRgTjRbvH/cOqRgTjRbvH/cBLARPqkYE40W7x/3DqkYE40W7x/3ASwGNba6kuVDDXUE7Kimmbqo5GLmjk4UMkAavEt/s+G7W+5Xuvio6Zu1qnrtuX81qJtuXkQ9MRXajsNjrLxcH6impIXSyKm6qJvJyquSJyqUr0hYxu2Nb/LdLnKqMRVSmp0d7SBm81qcPCu+p1dJ0mLMerlfS+jno7Xi9yZqnc26eWfj3R/vqdYxp0Q9ZK99PhK1sgj3Eqq1NU9eVGIuSfOq9w5RiDHeMb8534UxHcJmO3YmyrHH4jcm/qI2DT3NIuXPal6zoOBaBoMRFm1Ge2fXPjIu2uag9KeGaonZBTxSTSvXUsYxquc5eBETdOg4f0L6QLvG2VbSy3RO3HVsqRr87UzcnzocKLdVfsxm7ulado2iRnfuRT3zk50DuFH0OOIHNTpvENsiXfSKOST6UaZnsbazU/lbT58HSK+uZo0O9P5Wnq9K8IpnKb0eFX2cDPejq6ujl12jqp6eT86KRWL4UOx3HodcURNV1DebTVZe9k1car3Paqn6yDYk0YY6sDXSV2HqqSBu2s1NlOxE4V1GaondyOFWj3aPXNLt6PjmG6X/AE271M5/CZy/acn1ZdKWP7SrelsT10rU97VKk6KnB7dFXwE+w50Rd6gc2O/WSjrY9xZKZywv7uS6pF/UcOVFRcl2lAo0i7RyVGlYBhulR/5LNPfEZT4xlK5GENL+B8RqyGO6fg6qdtJBXokSqvAjs1avczz5CfIqORFRUVF20VD+fJLcFaRcXYRcxtqusjqVq7dJUfwkKpwI1fc/3VRTu2sRnkrh8diPoDTOdWhXMuyr7x/mPmu0Dj2AtPWHbwsdJiKJbJWO2tdVdXTuX426z59pOE67TzQ1EDJ6eWOaKRqOY9jkc1ycKKm6hsbd2i5GdMvgNOwzStAr3GkUTTP7T3TyS9AAZHRAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD8VURM12kA/QVZ0v9FQ63Xeps2ALfSVbYHrG+51ebo3uTaXWmNVM04HKuS8GW2vIavojdMM8qvbitsDc9pkVvpsk8Mar4VJmuT+ggKG2HontKtuna+ur7deI0X2zKqhYzNO7FqFLG6EuiAwzpDqY7NWwLZL89PaU0smqiqFTd1t+1mu/qVRF4M8lKjsoAAAAAAAAAAAAAAAAAAr90a/5K4f7+k+rKwU5Z/o1/wAlcP8Af0n1ZWCnOppD6TA+WGwg3DIaY8G4ZDTT3OV6/h/Nw/QAYmyAAAAAAAAAAAAAAAACM6R+scPfLfNcSYjOkfrHD3y3zXGWxzkNXjXuF3ubqxdZKDvaPzUMww7F1koO9o/NQzDhV7Uu7o3M0d0fQABxZwAAAAAAAAAAAAAAAAAAfLi/9p61UnyDPNQoA4v/AGnrVSfIM81DaYd+b5PNPxB5LH6v4soAG0eagAAAAAAAAAAAAAcD6LX8bhr4tV/pHfDT4iwzYcRLAt7tdPXLT6rWtdRV1Gqyzy7uSeACkoLidTTAfFig8VfSOppgPixQeKvpAp2C4nU0wHxYoPFX0jqaYD4sUHir6QKdmbYevlB3zH5yFt+ppgPixQeKvpPuHRxgeGVkseGqFr2ORzXI1dpU3F3QJWcZ6K78lrP367zFOzGrxDh+zYgp46e9W6Guiifq2NlTNGuyyzApCC4nU0wHxYoPFX0jqaYD4sUHir6QKdguJ1NMB8WKDxV9I6mmA+LFB4q+kCnYLidTTAfFig8VfSOppgPixQeKvpAkts620vyLPNQg/RD/AMk11+PB9cwn8bGxsaxiI1rURERN5DEvNrt95t0luulJHVUkior4pE9q7JUVP1ogFGgXE6mmA+LFB4q+kdTTAfFig8VfSBTsFxOppgPixQeKvpHU0wHxYoPFX0gU7BcTqaYD4sUHir6R1NMB8WKDxV9IDQ3/ACX2DvRPpUlxjWyhpLbQQ0FBAynpoW6mONm41OBDJA5P0VNXLT6LkhjVUbVXCGKTLfaiPf8ASxCqBdDTdhipxZo8rbbQM1dbE5tRTMz929i7be6rVcicqoUynilgmfDNG+OWNytex6ZOa5FyVFTeU02IUzFzN696B3rdWH1W6Z/qiqc478sp/wB2PgAHQfbrgaC8BWzC2FKK4up45bxXQNmnqHJm5iPTNI2rvIiKmeW6vzZdHK26NNPEVkw1BZ8SW6rrH0jGxU9RTK3VOjRMkR6OVNtEyTNN1N3b21k69EZhjesd4Xm/WN5a0mxTREROTxjFPR/GdI0u5crtzXMzPrzjLL4Zevk7Pg7WDiK9EbhvesF28aP1j5XojsP72HrovdfH6TJwuz0nQ1WxbqJ8Y+7uAOGr0R9i3sOXJf8AesPleiPs29hqvX/fs9BOF2ek5aq4v1M+Mfd0PHGjbCWL2PfcrayGscm1WUyJHMi8Kqm07+8ilcNJ+iPEGDEkrof/ANUtCLn01EzJ0SfCN973UzTubh0peiQtO9hitX/iW+g85OiPtb2OY7ClU9rkyVHVTclTg9yda9Oi3fjlL6PB7XpJhtUUxamqjozMftOfq+nYrkDf46uWH7ve3XDD1mls8MubpaV0qPY13CzJE1KLwb29tbSaA1VUZTlD0yzXVcoiqqmaZn4Tyx4ZwEpwLj7E+DahHWe4O6W1Wb6SbN8D+H2u8vKmS8pFgWmqaZziUv6Pa0iibd2mKqZ+E+tbfRvpow3ilYqG4qlmujskSKZ/8FKv6D+HkXJeDM6gfz4On6NtM+IsJwpQVzFvdta3KOGaVWyRcCNkyVdTyKi8mRsrGn/C54vO8Z9BuW7h8/pmfpM/SfFbgFfvZKR8TXf4l/6Q9kpFxOf/AIkn3Z2uG2Ol9XzGqGMdT/7U/dYEFf8A2SkXE5/+Ip92PZJw8T5P8RT7scNsdL6mqGMdT/7U/dYAHAPZJwcT5P8AEE+7Hsk6fifL/iCfdjhlnpfU1Rxjqf8A2p+7v4OA+yTpuKEv+IJ92PZJ03FCby9Puxwyz0vqao4x1P8A7U/d34HAvZJ0vFGby9PUHskqTijN5enqDhlnpfVNUcY6n96fu76DgXskqTilP5cnqH77JKj4pT+XJ6heGWel9TVLGOp/en7u+A4H7JKi4pVHlqeofvskqHinUeWp6g4ZZ6X1NUsY6n96fu72DgnskqHinU+Wt9Q/fZI0HFSp8sb6o4ZZ6RqljHU/vT93egcF9kjb+KlV5Y31T99kjbuKtV5W31Rwyz0k1Txfqf3p+7vIODeyRt3FWr8rb6p++yRtvFar8rb6o4ZZ6Rqni/Uz40/d3gHB/ZI2zitWeVN9UeyQtfFes8qb6o4ZZ6Rqni/Uz40/d3gHCPZIWvivWeUt9A9khauK9b5S30DhlnpGqmL9TPjT93dwcJ9khaeLFb5S30D2SFo4s13lDPQOGWekmqmL9TPjT93dgcK9khaOLNd5Qz0D2SFn4s1/lDPQOGWekaqYv1M+Mfd3UHC/ZH2bi1X8+z0D2R9l4tXDn2DhdnpGqmL9TPjH3d0Bwz2R9l4t3DnmD2R9k4t3DnmDhdnpGquL9TPjH3dzBwz2R9j4t3HnmH77I+x8XLjzrBwuz0k1VxfqZ8Y+7uQOSYO052XEmJqGxxWWvp5KyTW2SPexzWrkqpmidw62Zrdym5GdM5tZp2HaToFcUaRRuZmM/l8gAHN0gAh2l7HtDo3wguJLjQ1NbAlQyDWoFajs3Z5Lt7WW0BMQVr9l5hTipe+ci9I9l5hTipe+ci9IzFlCB9EHdqix6FsVXClerJkoHQsei5K1ZFSPNOVNXmcn9l5hTipe+ci9JodIfRL4Lxjgm74YqsM3yGO40zoklR8S627da7LVbeTkRct/IgqoXB0I9Dlga76OLRfsTNrbhX3SmbVZMqVijha9M2tajdtVyVM1VV28yny5Z7W4WR0KdEvDg7A9JhjEViq7glAmt0lTTStRyxZqqMe135u4iou5ltbWalfWn/obWYVsVRijBNVV1lBStWSsoahUfLFGm7IxyImqam6qKmaJmualcKaeamqIqmnlfDNE9HxyMcrXMci5oqKm4qLvlvK3ousNSMdC3Bdynhe1WvbLURoiou0qZZLmmRUi6yUkt0q5bfA+CjfO91PE92qcyNXLqWqu+qJkmYH9Dehzx7JpD0YUV3rHItzpnrR1+SZaqViIury/Sa5ruDNVTeOjlCOhr0zUmixt6pbpbqy4UdwWKSNlO5qLHIzVIq+2XfRU8VC3ehbSbbdKFhrbvbLdV0EdJVdLOZUK1Vcupa7NNSu57YqJ4AAAAAAAAAAAAAAACv3Rr/krh/v6T6srBTln+jX/ACVw/wB/SfVlYKc6mkPpMD5YbCDcMhpjwbhkNNPc5Xr+H83D9ABibIAAAAAAAAAAAAAAAAIzpH6xw98t81xJiM6R+scPfLfNcZbHOQ1eNe4Xe5urF1koO9o/NQzDDsXWSg72j81DMOFXtS7ujczR3R9AAHFnAAAAAAAAAAAAAAAAAAB8uL/2nrVSfIM81CgDi/8AaetVJ8gzzUNph35vk80/EHksfq/iygAbR5qAAAAAAAAAAAAABzrTNpDrMCPtTaS3QVnTqSq7XXq3U6jUZZZfG/UdFON9Enhi/wCIpLCtktc9clOlRrutIntNVreWefDkvgAjfshLxxdoOeePZCXji7Qc88gfUzx5xYrvA30jqZ484sV3gb6QJ57IS8cXaDnnj2Ql44u0HPPIH1M8ecWK7wN9I6mePOLFd4G+kCeeyEvHF2g555727T7d6q4U9M7D9C1JZWsVUmdtZqiHPOpnjzixXeBvpMuz6N8cw3ejlkw1WtYydjnOVE2kRyZrugW5ILphxxVYGtFFW0tDDWOqahYlbK9WoialVz2u4To5Z0RuH7ziHD1sp7Lb5q2WKrV72x5ZtbqFTPbAhXshLxxdoOeePZCXji7Qc88gfUzx5xYrvA30jqZ484sV3gb6QJ57IS8cXaDnnj2Ql44u0HPPIH1M8ecWK7wN9I6mePOLFd4G+kCeeyEvHF2g5549kJeOLtBzzyB9TPHnFiu8DfSOpnjzixXeBvpAt/SSrNSwzKiIsjGuVE3s0zI/pMxHNhPBtZfaemjqZKd0aJG9yoi6p7W7qd031vY6Ogp2PRWubE1FRd5ckIjpttdwvOje42610slVVyPhVkTN12UrVX9SKByz2Ql44u0HPPHshLxxdoOeeQPqZ484sV3gb6R1M8ecWK7wN9IE89kJeOLtBzzx7IS8cXaDnnkD6mePOLFd4G+kdTPHnFiu8DfSBPPZCXji7Qc88eyEvHF2g555A+pnjzixXeBvpHUzx5xYrvA30gWqwPeZcQ4Stt6mhZBJVwpI6NiqqN212kzN0RrRdQ1ds0f2agr4H09VBTI2SN+61c12lJKAKI44XPGt9XhuNR9Y4vcURxztY1vqf1lUfWONbiXs0vQ/w95+/wB0fWWmABqXqTPslnut8rFo7Pb6mvqEYsixQRq9yNRURVyTe208Ju00c48X+aV48lcTPoUP5Taj+zJfPjLVGw0bQ6btG6mXwvpD6V38L0zg9u3ExlE+vP4qRpo3x6v80rv5Op9Jo1x6v807rzCl2gZ+LqNstF/8gaZ1VP7/AHUmTRnj5f5p3TmT9TRjj9f5qXLm/wB5bzEOL8L4fVW3m/UFHIm3rT5k1zxE9t+oh1dp00eU7lbFcKyry34aR6J+0jThVodin2q/o72j+leM6TG6s6Luo2xFWXirsmi/SAv81Lj4iek/U0W6QV/mrcPFb6TvadEBgNVy1u8Jy9LN9Y94tPOj56+2qLjH8akX7FU4cH0bpuzOP+kMf/T/AGq+6s2JsKYiw02B19tNRQJUK5Illy9vqcs8sl3s08JpDsnRHY5wzjKksaYfrn1LqV86zNdA+NWo5GZe6REX3K7hxs6d6imiuYpnOH12EaTpGlaJTd0mjcVznnGUxl65iOX18gADE2SZU2i7H9TTRVEGGat8UrEexyOZk5qpmi+64D06lGkTitWeMz1i3uEvyUtHeMP1aG0NvGH25jPOXlN30906iuaYt0+qe37qXdSjSJxWrPGZ6w6lGkTitWeMz1i6ILxdb2yx6/6d1dH7/dS7qUaROK1Z4zPWMS66OMc2uifWVuGbgyCNM3vaxH6lOFUaqqicpdwCcOo2y5U/iBpuf9VqnL5/d/PgHd+ikwNQWtaXFtqp2U7aqfWK2KNMmrIqK5siJvKupci8K5LuqufCDV3rU2q5pl6PhWJWsS0WnSLfqifhsn4wGfYrPcr7cG2+00rqqrciubE1yI5yJu5Iq7fcQwD1pKiejqoqqlmfDPC9HxyMdk5jkXNFRd5ThGWfrd65u9zO45fhnyJYui/SAn81Lj4iek+V0Y4/T+alz5stLocxhs1wRTXObUpXQuWnrGt2k11qJ7ZE4HIqO+dU3iZm1o0C3XTFUTPreY6T6b4hot6qzdtUxVTOU8v3UlXRrj5P5p3XmFPldHGPE/mld/JnF3AXi6jbLD/8gaZ1VP7/AHUhXR3jtP5o3nyR/oPldH2OU/mje/In+gvARvGeOMMYQgR98ukUErkzZTs9vM/uMTby5VyTlJVoFumM5qZrHp1p9+uLdqxFVU/CM5lUFcA44T+aF98gk9B8rgTG6fzPv/8Ah0vqnY8QdEexHujsGG3Ob72WtmyVf7jfWIrU9EFjqVyrHT2aBN5GUz1856nUqt6PH5p8H09jTcfuxnOj0U99X2zQRcD41TdwfiD/AA2b1T5XBWMk3cJX5P8Al0vqk9puiCxzE5Flp7NOm+j6Z6ea9CVYf6I+NXtZfsNuY330tFNqv2HZecKbejz+aY+S39Ox61GcaNTV3VffJxZcHYuTdwrfE/5fL6p8rhHFabuGL0n/AAEvqlxsGY8wri5mVku0Us+WbqaT2kzeH2i7apypmnKSY7VOH0VRnFT5u96daZo9c27ujxTMfCZmJ+iiC4VxQm7hu8J/wMnqnyuGcSJu4euyf8HJ6C+JFcQaRcEWGV0NzxJQxyt2nRxOWZ7V4FaxFVPnJVoFFMZzVk5WPTnTL9W5taNup2RMz9IU2XDmIU3bDdE/4ST0HjVWe70sLp6m1V0ETfdPkp3tandVULYt036OFk1K3qZE/OWimy83M1Wl3GmFcS6IL+yyX2jrJdbiXWkfqZMtej29Q7J2XzGKrRbW5maa88mzsek2JTet272hzTFUxGc5+rOctiqoAOg+4D9a1zs9S1Vy3ckPwsJ0HPucU92k/wBYy2LW+1xRnk1mMYjxbodelbndbnL1Z5csxHLlO1XtdrdB/QZURUyVEXulb+i+a1t7sGpaif8AZpc8k/Sadm/oW9UTVus/k+cwb0x4z0unRt53Oefr3WfJGfJuYcJAB0X2wAXv2MYbc1NXh60u2t+jjX7Ds6Po0388pyyfO4/6Q0YNve6omrd5/HLLLL7qhaEkz0r4dROy081S6pq6LDuH6GrbV0VitdNUMz1MsNJGx7c0yXJUTNNo2httGsTZpmJnN5b6R43Ri9+m7RRucoy9c5/GZAAdl88Ec0h4MsWPMPLYcRQzS0SzNm1MUqxu1Tc8ttO6pIzT40xBRYUwrccR3GKolpLfAs8rIGosjmpvNRVRM+6qAct9jDon7XXLy95gX7oV9G9Za5oLW66WysVq6zUJUrKjXb2qa7dThRMl5UMT2W+jntJivyWn++MXEfRbYNhtErsP2O9VVxcxdZZVxxxRNdvK9WvcuScCJt8KbpFU7vVvntN4rbXVanX6Ookp5dSuaapjlauXJmh2joUsE4B0h191sGKLdVPuVNElXTTQ1To0fFmjXtVE2s0VzVRd/VLwHE7hVz19fUV1U9ZKioldLK9ffOcqqq+FTuHQZVdtsONr1iy/XCC22ihtvS0lTO7UsSWaVmobnwqkb1+YDvFT0MmiiOmle23XLVNYqp/29/AUQP6NVemXRa6lma3HNmVVY5ETXt3a7h/OUEO/dCpojwnpKtF9qsSLcEkoaiKOHpadI0yc1yrnm1c9xDk2lOyUWG9I+ILBbtd6ToK+Wnh1x2qdqWuyTNd9SzP/AMP/APJ3FffdP5jyvOn3+WrGH9rT+coEh6FnAuH9IOkSssmJIZ5aOK1yVLWwyrGurbLE1NtN7J7touxoz0fYb0d2qptmGYaiKnqZ9fkSaZZFV+pRu6vIiFTugQ/liuP9hTfXwF3RBIACoAAAAAAAAAAAAAK/dGv+SuH+/pPqysFOWf6Nf8lcP9/SfVlYKc6mkPpMD5YbCDcMhpjwbhkNNPc5Xr+H83D9ABibIAAAAAAAAAAAAAAAAIzpH6xw98t81xJiM6R+scPfLfNcZbHOQ1eNe4Xe5urF1koO9o/NQzDDsXWSg72j81DMOFXtS7ujczR3R9AAHFnAAAAAAAAAAAAAAAAAAB8uL/2nrVSfIM81CgDi/wDaetVJ8gzzUNph35vk80/EHksfq/iygAbR5qAAAAAAAAAAAAAABxvoksT3/DkliSyXSehSoSo13W8vb6nW8s803s18IHZAU86puPeM9b+z6B1Tce8Z639n0AXDBTzqm494z1v7PoHVNx7xnrf2fQBcMFPOqbj3jPW/s+gy7PpJx1NdqOKTEta5j52Nc1dTtorkzTcAtwAct6IzEN6w9h62VFluE1FLLVqx7o8s3N1CrltgdSBTzqm494z1v7PoHVNx7xnrf2fQBcMFPOqbj3jPW/s+gdU3HvGet/Z9AFwwU86puPeM9b+z6B1Tce8Z639n0AXDB4W97pKCnkeubnRNVV4VVEIjpsutwsuje43G11UlLVxPhRkrMs25ytRd3kVQJqCnnVNx7xnrf2fQOqbj3jPW/s+gC4YKedU3HvGet/Z9A6puPeM9b+z6ALhgp51Tce8Z639n0Dqm494z1v7PoAuGCNaL6+ruej+zV9fO6oqp6ZHSyO3XLmu2pJQBRLHm1jm/p/WdT9a4vaUUx+mWPMQJ/WlT9a41uJezS9D/AA+5+93R9WjABqXqTrnQo/ynTf2bL58Zasql0Kf8p8v9my+cwtW9zWMc97ka1qZuVVyRE4Td6BzXzeN+nH90/TH+WsxViC04Yss13vNU2npYtrhc9281qb7l4PsKv6R9NWJcSTS0tolkstrVVRrIXZTSJwvem2ncbkm9t7pp9NGO6jG2K5ZIpXJaaRzo6GLcTU78ip+c7LPkTJN4gp09K0yqudzRPqfXejnopZ0S1Tf0qndXJ9eU8lPy27Z+Hw7f17nPcrnOVzlXNVVc1VT8B+plmmaKqb+SnQfbPwFktF2AdEOMMPsq6CkrKipiajaqGorXpNE7lRitRUXeVEyXu5okmqtBOjyZqpHQVtOvDHWPVf2szu06DcqjdRMPkL/ppoOj3Zs3aK4mOXOI+6o4OraedGtmwHDbKi0VldM2tkka5lS5rtRqUau0rWpwnKTq3LdVurc1cr6PQNOs6fYp0izP9M/4nIABwdxfPCP5J2fvGD6tptDVYP8AySs/eEH1bTlXRA6RMTYKxNaILHUU7aeWmdLNFLAj2yLqskzX3SJtbypun0Vd2LVvdVPz/ouHXcR02dHs5bqZnl7M5dqBwLDvRHUjmtZiHD00TvfS0MiPRf7j8svGUnNt01aOq1qaq+OpHr7yoppGqnzo1W/rJTpVqrkqZtJ9G8U0acq7Mz3f1fTN0QELk0qaPWR6t2KqFU/R1Sr4ETMiWK+iAwnb4HssMFVeKnL2i6hYYUXlVyarwN+dC1aRapjOaoYbGB4jpFW5os1fOJiPGcofHRZXOmp8BUdsc9q1NZXNexm/qGNcrneFWp85Vw3mNcU3jF98ku96qEkmcmpjY1MmRM3mtTeT9a7q5mjNJpN7fbk1Q9l9HsLqwvQabFc51euZ2ZyAH3BDLUTxwQRPllkcjGMY1Vc5y7SIiJuqYG6mcvXKwnQdvmWnxNGqrrKPpnN4NUqS5/qRP1FgCA6C8GS4MwRHTVrUS5Vj+mKtE29QqoiNZn+iieFVJ8fQ6NRNFqIl4N6R6Xb0vE7t2166ZnKO3KIjP55ABF9KOKY8HYJrr0updOxut0rHbj5nbTU7ibq8iKZqqopiZlqtHsV6RdptW4zqqnKPmg+nbS0mFUfh/D72SXp7f4aZURzaRFTa2txXqm2iLtJurwFX6+sq6+slrK6plqamZ2qkllernPXhVV3T8rqqorqyasq5nzVE71klkeubnuVc1VfnPE+fv36r1Wc8j3XBcEsYVYiiiM6p5avjM/bZH+QAnuhjR7Pj2/SRzSPp7VRojquZie2XPcY3P3y5Lt7yIvIi4qKJrqimnlbHS9LtaHZqv3pypp5UCBdi26MsBUFI2miwrbJWomWqqIUmevLqn5qc50y6FbS6y1F8whSLSVlMxZJaJiqsczE211CL7lyJt5JtLlllmdyvQLlNOeeb5XQ/TjQNIvxaqpmmJ9UTOWXz9fqVxp55qadlRTzSQzRuRzJI3K1zVTcVFTbRSxmhDTL+EMrBjGqjZUMjVae4SKjUlRqZq2Te1WSbTt/f292t4OvZvVWqs6W+xbB9GxSzvd6PX8J+Mf7sdQ0t6XLziuunoLTUTW+xNVWMjjVWvqE/OkVNvJfzdzhzXbOXg+omPlkbFExz3vVGta1M1cq7iInCcLlyq5VnVLs6FoOj6DZi1Yp3NMf7nO2XyCeUOh/SPWUzaiLDMzGOTNEmqIonfO1zkVPnQ0OJ8G4owyiOvljrKONVySVzNVGq8Grbm3PkzE2q4jOYlLWJaHdr3u3dpmrZFUTPhm0IAODuhYXoOvcYp7tJ/rFeiw3QdfisUfGpP9Y7Whc/T/vwfM+mH9nvfp/7QsCVu6MDr7YO9pfOQsiVt6MDr/YO9ZfPQ2enczLzn0M/u9vuq/6y4UADRPbA/oMm4fz6buof0GNphv5vl/l5p+If/wBf9f8AEABtHmoAABoNIeHExdgi74ZWs6SS40zoNf1vXNbz39TmmfczQ34A/m3pv0f9TXHLsMpdFuaJTRzpULT6znq89rU6p25lu5kXwzY7piW+0tjstL01cKtysgh1xrNWqIq5ZuVETaRd1SzPR6YQqFq7Ljilhc+BIvwfWOan4tUcr4lXu6p6Z8jU30KxWK6V1kvVHeLZO6nraKdk8Eie9e1c0Xl3NzfIrseFehf0nXSsYy701BYabP28tRVMlcifotiV2a8iqndQnumrQtiK2YGw/gPRxh+qulFHM+4XaudLEx1TU6nUMVUc5NxqvyRNpEVN/NV9rJ0X9N0jG29YLm6ba1Ee+krE1t676o1zc2pyZr3T2pOizfdb/b7XbMEtgbV1cUCzVFw1Soj3o1V1LWJw8IHE36AtLrGOe7BlSjWpmq9NQbnjnMj+q9b/ABKf5N30H8qALff/AA//AMncV990/mPK86ff5asYf2tP5ylhv/h//k7ivvun8x5XnT7/AC1Yw/tafzlA6R0CH8sVx/sKb6+Au6Ui6BD+WK4/2FN9fAXdEJIACgAAAAAAAAAAAAAr90a/5K4f7+k+rKwU5Z/o1/yVw/39J9WVgpzqaQ+kwPlhsINwyGmPBuGQ009zlev4fzcP0AGJsgAAAAAAAAAAAAAAAAjOkfrHD3y3zXEmIzpH6xw98t81xlsc5DV417hd7m6sXWSg72j81DMMOxdZKDvaPzUMw4Ve1Lu6NzNHdH0AAcWcAAAAAAAAAAAAAAAAAAHy4v8A2nrVSfIM81CgDi/9p61UnyDPNQ2mHfm+TzT8QeSx+r+LKABtHmoAAAAAAAAAAAAAHOdM2jytx2+1OpLjT0fSSSo7XWK7VavUZZZfF/WdGAFdvY93njDQc08ex7vPGGg5p5YkAV29j3eeMNBzTx7Hu88YaDmnliQBXb2Pd54w0HNPMi3aArxS3CnqXX+gckUrXqiRP28lRSwIAEF0xYHq8c2iioqSugpHU1QsqulaqoqalUy2u6ToAV29j3eeMNBzTx7Hu88YaDmnliQBXb2Pd54w0HNPHse7zxhoOaeWJAFdvY93njDQc08ex7vPGGg5p5YkAeVJEsNLDCqoqxsa1VTfyTIj+kzDk+LMG1lip6mOmkqHRqkkiKrU1L2u3u4SUAV29j3eeMNBzTx7Hu88YaDmnliQBXb2Pd54w0HNPHse7zxhoOaeWJAFdvY93njDQc08ex7vPGGg5p5YkAaXA9mlw9hK22WaZk8lJCkbpGIqI5c120zN0AAKK6Qkyx/iJP61qfrXF6ii2kZMtIWI0/raq+tca3EvZpehfh97xe7o+rQgA1L1N1noVP5UH/2dN5zDumnm7vs2iq9TwuVs08baVip8I5Gu/ZVxwroVv5UV/s+b6WHV+ipR/UuTU7iXCHVdzJ/25G10ectFqmO15hjtqm76S2KKuSdx9ZVQAMi2MgkuVLHVO1MDpmJKue41VTP9Rqnp1U5Rm7zoS0M2+5Ye/D2LoHy9PQr0nSo5Wa3G5NqVVT3ypttTcRMl21Xa53pb0Z3XAlesqauss8z8qerRu5+hJwO/Uu9volyImMjjbHG1rWNREa1qZIiJuIh4XOho7nQTUFwpoqmlnYrJYpG5tcnKbuvQaJt7mOXa8c0X0z021ptV+5/VRV+XZHwy2TH7/FRXC9/u2GrzDdrNVvpqqJd1Ntr277XJuK1eBS2uiXSZacd0CRLqKO8wszqKNXbv6cfC39ab+8q8M006I6zCMsl4srZauxOdm730lJnvO4W8DvmXeVeY22urLbXw19BUy01VA9HxSxuyc1eFFOhbu3NFr3NXI+30/DdA9JdFi/Zq/q+FXxjsqj/cuWO2wvRhdasO/Lz+awrkdD0l6SJcdYVstLcabWrpQSya/IxMo5mua1Eciby7S5pucHAnPDHpVdNy5NVPI7/o3oV7QcPosXoyqpmr/tIADrt8vlg78kbN3hB9W0r70Xv5VWTvF3nqWBwb+SFm7wg+raR3SRoysGPKykq7tVXGnlpY1jYtLIxqOaq57eqa79WRv79uq5Z3NPY8NwTELOH4rv8Ae9mN1yevlzUwB3vG/Q8z01I+qwndJKx7Ez6UrEa17/ivTJM+RUTunCq2lqaKrlpKyCSnqIXKySKRqtcxybqKi7imlu2a7U5VQ9gw7F9ExKmatGrzy5Y5Jj5PEAGJsgzrXZ7tdZEjtdrra56rllTwOkX9lFPG3VtVbq2Ktop3wVETtUx7d1F/97xajQPpRbjGmWy3dIYL1Ts1TVYiNZUsTdc1N5yb6J3U2s0TsaPaou1bmqcmjx3E9Jw2xv8AZtbuI5fXll25ZeuPm49hbQdjm8PY+tpYbPTLuyVb01eXIxua58i5HedGuijDeCnNrI2uuN1RMlrJ2pmzh1tu4z9a8pPwbe1olu1OceuXlOKelOIYjTNuqrc0z8KfVn3zyz45dgADsvnArv0X13etTY7Cx+TGsfWSt4VVdQxfmyf4SxBVLoq5HP0nxtduR26Jre5qnr9KqdPTqsrMvrPQuzTcxWmZ/LEz+2X+XJQAaN7SFwuh1skdn0WW6TUIk9wV1ZKuW7qlyb+wjSnpejRw1rdHmG2s9ylqpcuaabDDqYmuZ7HwXp/eqp0O3bjkqq9fyj/+t+ADcPJ1EcdUkNvxvfaGnajYaa5VEUbU3mtkciJ4ENMSHSZ/KRif+16v65xHj5mv2pfo3RKpq0eiZ2R9A6r0LdtprhpQ16oja9aGhkqYkcmaI/VMYi/Nq1U5Udi6Ej+Ui4f2RJ9dCZdGjO7S1vpFXNGF35pn8srSnP8Aoh/5HL98WD6+M6Ac/wCiJ/kcv3cg/wAxGby/zdXdLxbBv7jo/wDzp/7QpwAD5x+gwsP0HX4jE/xqX6JSvBYfoO/4tif49L9Ep2tC56P9+D5n0w/s979P/aFgCtnRf/lDYe9JPPQsmVs6L/8AKOxd6SeebPTuZl516Gf3e33VfSXCwAaJ7W+o9uRqcqH9BT+fcO3MxP0kP6CG0w383y/y80/EPl0f9f8AEABtHmoAAAAAw7zbLfebXU2u60cNZRVLFjmgmbqmPau8qFcsa9CRYq2qkqcKYkqbSxyqqUtXD0xG3ka/NrkTu6peUsyAKZSdCJjJJMo8UWFzOFySovg1K/SSHBvQm3C3Xygul0xnSp0nURz61TUTn6tWOR2Wqc5MtzdyUtYBkZviViSRPjVVRHNVq5cpXf2I2BeMeI/Hh+7LFgDn+hnRXZdFtDcaSy3C4VjK+Vkki1bmKrVaiomWpanCQrGXQx4OxRiq54irL7foai41L6iWOJ8Woa5y5qiZsVcu6p3UAcn0QaCsNaM8Tz3+z3a7VdRNRupHMqnRqxGuexyqmpYi55sTf31OsAAAAAAAAAAAAAAAAAAV+6Nf8lcP9/SfVlYKcs/0a/5K4f7+k+rKwU51NIfSYHyw2EG4ZDTHg3DIaae5yvX8P5uH6ADE2QAAAAAAAAAAAAAAAARnSP1jh75b5riTEZ0j9Y4e+W+a4y2Ochq8a9wu9zdWLrJQd7R+ahmGHYuslB3tH5qGYcKval3dG5mjuj6AAOLOAAAAAAAAAAAAAAAAAAD5cX/tPWqk+QZ5qFAHF/7T1qpPkGeahtMO/N8nmn4g8lj9X8WUADaPNQAAAAAAAAAAAAAOC9Fm97JcN6h7m5tqc8ly/ojvRENIuj6z45dQuutVXwLRJIkfSz2Nz1epzz1TV/NQCoGvzf00njKNfm/ppPGUsj1AMI9tL5z0X3Y6gGEe2l856L7sCt2vzf00njKNfm/ppPGUsj1AMI9tL5z0X3Y6gGEe2l856L7sCt2vzf00njKZlimm/DdB/CyfxmP3y/nIWF6gGEe2l856L7s9aTQPhOmq4allzvavika9qLLFkqoue3/BgdYONdFW97ML2hWOc1enV3Fy94p2UjGkPBNrxvb6aiulTWQR08qysWmc1qquWW3qmrtbYFONfm/ppPGUa/N/TSeMpZHqAYR7aXznovux1AMI9tL5z0X3YFbtfm/ppPGUa/N/TSeMpZHqAYR7aXznovux1AMI9tL5z0X3YFbtfm/ppPGUa/N/TSeMpZHqAYR7aXznovux1AMI9tL5z0X3YHVLZ1tpfkWfQhCOiFc5uie6q1ytXVwbaL8MwnkEbYYI4WqqtY1Goq7uSJkarGmHaPFeHaix3CWoip51YrnwKiPTUuRyZKqKm6nABSrX5v6aTxlGvzf00njKWR6gGEe2l856L7sdQDCPbS+c9F92BW7X5v6aTxlGvzf00njKWR6gGEe2l856L7sdQDCPbS+c9F92BW7X5v6aTxlGvzf00njKWR6gGEe2l856L7sdQDCPbS+c9F92BMNDqq7RjYFcqqq0iba91SWmtwzZ6awWGjs1HJLJBSR62x0qor1TlyRE/UbIAUY0kplpFxKn9bVX1zi85RnSZtaR8Tf2vV/XONbiPs0vQfw+95vf8Y+qPAA1L1R1foV/5Uv+Am+lp3rTlZpL5otvVJCxXzxQpUxom6qxuR6onKrUVPnOCdCz/Km3vGb/AKS2CoioqKiKi7qKbjQqYrsTTPxzeSel+kVaNjdF6nlpimfCZfz5B0XTno/qMGYlkqaSBy2SukV9LIibUSrtrEvAqb3CncXLnRqq6JoqmmXqOhaZa02xTfsznTV/uXfC1OgHSfR4htNNhy81LYb1SsSOJ0jskq2ImSKi/nom6m6u6m/l14/n1G98cjZI3OY9qorXNXJUVN9DruA9POJLJDHR32Bt8pWJkkj36ioanx8lR3zpnymx0fToiNzc8Xn2P+hVyu5Vf0H4+uaeTwnk+U8i0ssbJYnxSsa+N7Va5rkzRyLuoqb6Fb9N2hmS26/iLCFO6Si231NAxM3QcLo+FnC3dTe2tyWt6IrByw5utF+STL3KRRKmfd1z7DR3vokGalWWXDLldvSVdRkif3Gp/wBRmv3dHuU5VS1eCYbj+HaRu7FqcvjEzERMePhMK9A2OJbs++XupustFR0clS/VvipI1ZGi76o1VXLM1xpp5fU9comqaYmqMp2AAI5r4YM/I+y/2fB9W025qMF/kdZP7Pg+rabc+mp9mH5w0nnq++fqHLNO+jCHGFtfeLRCyO/UzNrLaSqYnvHfpfmr8y7W51MEuW6blO5qZtA069oF+m/ZnKqP37J7Jfz7mjkhmfDNG6ORjla9jkyVqptKipvKfBZfohdFa3dk2LMOU2dxY3VVtLGm3UNT37U/PTfT3ycu7Wg0F+zVZq3Mvc8Hxeziujxet8vxj4xP+8k/EMyy3Kts91prpbp3QVdNIkkUjd5U+lN5U30MMGGJy9baVUxXTNNUZxK7+jHGNFjbCsF3ptTHOn8HVwIu3DKibadxd1F4F4cyUFLdD+OqnAuKG1i6uW21OUddA33zM9pyJ+c3dT503y5Frr6O6W6C42+ojqaWoYj4pWLmjmqb7RdIi9T6+WHiPpNgVWFaTnRH/jq9mdnZ8v3j5soAHafNBWXoubZJDi+03dGrrVVRLDn+nG9VX9UjfAWaIRpqwauNMET0NO1v4RpndMUSrtZvRFzZnwORVTu5LvHX0q3Ny1MRyt96NYhToGI27tc5Uz6p7p+Pyn1qXg+54paeeSCeN8UsblY9j0yc1yLkqKm8p8Hz73aJz9cBc7QRdo7voqskjXIr6aDpSRN9qxrqURf7qNX5ymJ0/QJpIZgm6zW+6q91lrnIsitTNaeTcSRE30y2lRNvaRU3Ml7eh3ot3PXyS+X9LcKuYjoOVqM6qJziNu2P8/JbgGLa7hQ3ShjrrbWQVdNKmbJYXo5q/OhlG95XilVM0zlVGUqg6Q9HuNqrHmIKumwzcp6eouVRNFJHCrmvY6RzmqipwoqGgXRzjxP5pXjyVxd0w6i6WynrYaGouNJDVTrlFA+ZrZJF4GtVc1+Y19WH0TOcy+70f060yiim3TapnKO34KWLo8x0n80b15I/0HVehiwniWyY5r668WOvt9OtsfC2SphdGjnrLGqIme7tNXwFigcreg026oqieR1sQ9NNJ03Rq9Hqt0xFUZZ+sOfdEV/I3fu5T/5iM6Cc96Iv+Rq/f8P/AJiI7N/mqu6Xz2C/3HR/+dP/AGhToAHzj9BhYjoO/wCK4m+PTfRKV3LE9B3/ABPEvylN9Eh2tC56P9+D5j0x/s939P8A2h34rX0X35SWLvN/nllCtXRfflNY+8n+ebPTuZl536F/3e33VfSXDQAaJ7W9KbbqI0/TT6T+gZ/P2k26qFP02/Sf0CNphv5vl/l5n+IfLo/6v4gANo82AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABX7o1/yVw/39J9WVgpyz/Rr/krh/v6T6srBTnU0h9JgfLDYQbhkNMeDcMhpp7nK9fw/m4foAMTZAAAAAAAAAAAAAAAABGdI/WOHvlvmuJMRnSP1jh75b5rjLY5yGrxr3C73N1YuslB3tH5qGYYdi6yUHe0fmoZhwq9qXd0bmaO6PoAA4s4AAAAAAAAAAAAAAAAAAPlxf8AtPWqk+QZ5qFAHF/7T1qpPkGeahtMO/N8nmn4g8lj9X8WUADaPNQAAAAAAAAAAAAANNiTFFgw4sCXy6QUK1Gq1rXM/b6nLPLJN7NPCbk4H0Wv43DXxar/AEgOmdU7AXGai/a9A6p2AuM1F+16CnoAuF1TsBcZqL9r0DqnYC4zUX7XoKegC4XVOwFxmov2vQfcOkrAs0zIosSUbnvcjWtTVbartIm4U6M2w9fKDvmPzkAvKavEWIbLh2niqL1cIqKKV+oY6TPJzss8tpDaHGeiu/Jaz9+u8xQJt1TsBcZqL9r0DqnYC4zUX7XoKegC4XVOwFxmov2vQOqdgLjNRftegp6ALhdU7AXGai/a9A6p2AuM1F+16CnoAvhG9skbZGLm1yIqLwopiXu62+y22S43SqZS0kSoj5X55JmqIm5yqh6WzrbS/Is81CD9EP8AyTXX48H1zANj1TsBcZqL9r0DqnYC4zUX7XoKegC4XVOwFxmov2vQOqdgLjNRftegp6ALhdU7AXGai/a9A6p2AuM1F+16CnoAvTa6+kudvhr6CdtRSzt1UUjdxycJkkR0N/yX2DvRPpUlwAo1pO2tJOJv7XqvrXF5SjmlFMtJWJv7WqfrXGuxH2aXoP4fe83v+MfVGwAah6o6r0LX8qjO8pv+ktkVM6Fz+VWLvKb6ELZm6w/mvm8d9Ov7nH/GPrLCvdqt17tc9sutJFV0k7dTJFImaLy8ipvKm2hXTSFoBu1DNJWYQmS40qqq9KTPRs7ORFXJr08C90swDsXtHoux/VDR4VjmmYXVnYq9U8sT64n/AHbCgt3tN0s9StNdbdV0Mye8qInMVe5mm2YR/QGrpqarhWCqp4p4nbrJWI5q/MpGq/RxgStcrp8KWpFXdWKnSPPxcjoVYdP5an3Gj/iDamP/AD2Zieyc/rl9VIwXMXRBo4Vc9jEHPy+uZNJot0fUzkdHhW3uVP6VqyJ4HKpw4uubYdufT/QcvVbr/b7qVAsJ0V1st1rw/h6C2UFLRQ9MTfwdPC2NvuW7zUQr2dS9a3quac31GEYlGJ6JTpNNO5ic/Vy8k5AAMTZr34J/Iyyf2dT/AFbTcGnwR+Rdj/s6n+rabg+mo9mH5x0nnq++fqAA5MAcK06aHPwm+fEuEqdErVzfV0LEySdd98afn8Ld/e293uoMd21Tdp3NTY4ZimkYbfi9Yn1/GPhMbJfz6kY+OR0cjHMe1Va5rkyVFTdRUPkt5pY0RWbGaSXGiVlsvWX49rf4OdeCRqecm33csisOMsH4hwjXLSXy3S0+a5RzImqil5WvTaXubqb6IaO/o1dmfXybXsuDekWiYrTEUTua/jTPL8tsf7OTQk50YaTr/gWZYaZW1tskdqpKKZyo3PfcxfeO8KLvopBgYaK6qJzpn1tvpWiWdLtTav0xVTPwlae3dEJguemR9ZR3akmy9szWWvTPkVHbfzohg1vREWZ9fT0tosFdUNlmax0lTI2LJFVEVURuqz/UVmMuzdd6PvhnnIdvh16fVm+Y1LwqjOvczPZnOX3/AHX8ABu3jLkWmrRBT4sdJfLBrVLess5Y3e1jqsuFfev5dxd/hSsV6tVystxlt12op6OqiXJ8UrdSvd5U5U2lL9GmxThiwYooulL7a4K2NPcK9Mns5WuT2zfmU6OkaFTcndU+qX2mBemF7QKYsaRG7txybY+8dk+KiILFYq6HOnke+bDN9dBntpT1zNU3x27aJ/dXunOrxoV0h25ztTZmV0ae/pahjkX5lVHfqNbXot2jlpeiaJ6S4XpUf03oidlX9P1/whVlvd5ssyzWi61tA9fdLTzuj1XdyXb+ck0elfSGxmoTFNYqfpNYq+FW5mrqsC41plVJsJ3xuW+lDI5PCiZGMmE8UuXJMNXlV4EoZfVOETdp9UZw7dynDtIndXIoq7Z3Msq548xpcmqysxRdpGLusSqc1q/3WqiHvonc5+lHDj3uVzlucKqqrmq+3Q+KPR7jmrVEhwleUz35KR8aeFyIT7RZojxzR41s94uVrjoaSkq2TyOlqGK5WtXPJGtVVz2stvI526LtVcTMTLpadpmHaNotyimuinOmfVExHw2QtEAD6B4SHPeiM/kavv8Aw/8AmIjoRzzojf5G77/w/wDmIzFf5qruls8F/uOj/wDOn/tCnYAPnH6CCxXQefxHEvytN9EhXUsX0HnW/EnytP8ARIdvQuej/fg+Y9Mf7Pd/T/2h3wrT0X35UWTvJ/nqWWK0dF7+VVk7xd56my07mZeeehf92o7qvpLh4ANE9qe1F/HYPlG/Sf0BP5/2/br6dPhW/Sh/QA2uG8lXyeZ/iH7Wj/q/iAA2bzYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFfujX/JXD/f0n1ZWCnLP9Gv+SuH+/pPqysFOdTSH0mB8sNhBuGQ0x4NwyGmnucr1/D+bh+gAxNkAAAAAAAAAAAAAAAAEZ0j9Y4e+W+a4kxGdI/WOHvlvmuMtjnIavGvcLvc3Vi6yUHe0fmoZhh2LrJQd7R+ahmHCr2pd3RuZo7o+gADizgAAAAAAAAAAAAAAAAAA+XF/wC09aqT5BnmoUAcX/tPWqk+QZ5qG0w783yeafiDyWP1fxZQANo81AAAAAAAAAAAAAA1N/w5Yr+sK3q1UtesGq1rXmarUarLPLu5J4DbEM0laQrbgV1A24UNXVdOpIrNY1PtdRqc881T85AMnqdYG4r2zmR1OsDcV7ZzJBvZA4b7SXb/AMv1h7IHDfaS7f8Al+sBOep1gbivbOZHU6wNxXtnMkG9kDhvtJdv/L9YeyBw32ku3/l+sBOep1gbivbOZPqLR9gmKVkseGba17HI5rki20VNxSCeyBw32ku3/l+setHp6w7VVcNMyzXVrpZGsRV1vJFVcvzgOvGtv1is9+gjgvNup66KJ2rY2ZuqRq5ZZobIARXqdYG4r2zmR1OsDcV7ZzJCNIunBmEcbV2GWYUqbi+jZE587axkaLrjEciIipymg9ki7iFWf4lH6oHVup1gbivbOZHU6wNxXtnMnKk6JHhwHW/4jF6D99kgnEOu/wAQi9AHVOp1gbivbOZHU6wNxXtnMnLE6JBm/gSv/wAQi9B++yQj4i3Dy+IDvTGtYxrGIjWtTJETeQxbvbKC70ElBc6SKrpZFRXxStza7Jc02u6iHD/ZIRcRbl5dCfvskIOItz8thA6j1OsDcV7ZzI6nWBuK9s5k5f7JCn38DXTy2H0j2SFNxGuvlkHpA6h1OsDcV7ZzI6nWBuK9s5k5h7JCk4jXfyuD1h7JCj4j3jyqD1gOn9TrA3Fe2cyOp1gbivbOZOY+yQouI958qp/WHskKDiPe/Kaf1wO4W6ipLdQxUNDTx09NC3UxxMTJrU4EMg4R7JG38R755RT+uPZI27fwPfufpvvAO7nKMR6CcKXy/V14nuV5hmrZ3zyMjlj1COcua5ZsVcs14TQeyRtnEi/89TfeH6nRI2rfwTiH5pab7w4V26LkZVRm7mhYhpOg1TVo9c0zPLkzvY6YQ7cX3nIvuz89jphHtzfOci9QxE6JGye+wZiRO4tMv+qfSdElh732D8Up3G0y/wCsYuC2ei2Os2K9fP7fZK9HuiHD+C8Qfhu319yqahInRNbUPZqUR2Wa+1ai57R0Y4gnRJYY99hLFqdyGmX/AFz7ToksJ7+FcYJ/w1N9+ZaKKaIypjJrNL02/ptzfL9W6q5M5dsBxROiRwhv4axcn/CQffH2nRH4NXdsGK07tFF96c3VdoBxpOiNwUu7ZsUt7tAz7JD7b0RWBl3bZiZO7b09cDsQOPp0RGBF3aLEad22r6T7TohsA78GIE/5Y8CV6UdH9ux9Q0VNX1tVSdKSOkY6FGrnqkyVFRU5EID7HHD/ABhufNx+g2qdEJo+3235P+VS+g+k6IPR1vyXxvdtE/qmGvR7Vc7qqPW2+iY9iGh2otWLs00x8PV9mn9jjYeMVz5tnoPz2ONi4x3LmmG7TogtGy7tXeG920VHqH2nRAaMl3bnc2920VX3Zw4JZ6Ls604v10+EfZ0u1UcVutlLb4Fc6KlhZCxXLtq1rURM+XaMk5amn7Rdv3qvb3bPV/dH0mn3RVv4iqm92z1n3R2Y9TQVVTVMzPK6gDmKae9FK/zmmTu2qsT/AEj6TTxopX+dLk7ttqk/0gjpgObJp10VL/Otid2hqE/0z7TTjorX+d1Ondpp0/6AOjGPX0dHcKR9JX0sFVTyJk+KaNHscnKi7SkETTZosX+eNH88Uqf9J9ppp0Wr/PO3p3Uen/SFiZpnOGsxHoIwNdZHTUcdZaJF28qWXNir8V6Ll3EyIxJ0N1Ar848V1LW8DqNqr4dUhPE0zaLl/nrak7r3J9h9Jpj0XL/Pmyp3Z8vpOvVotmqc5pbyz6TYrZp3NN6cu3KfrEodb+hzw1G5Frr7dahE3okjiz8KOJjh7RDgGyzx1EFkSpqInI5stVK6VUVNtF1Krqf1HqzS5owfuY8w+ndrWJ9KnszSno1fuY+wwndukKfS4tOj2qeSli0j0gxPSIyuXqsuycvpkmIIozSTo6f7nH2FXdy70/rnuzH2BH+4xrht3cukK/8AUZ2nSQGiZjLCEnuMVWJ3cuES/wDUe7MTYbf7jEFpd3KyNftA2wNey92Z/uLvQO7lSxftPdlfQv8AcVtM7uStX7QMkHwySN/uHtd3FzPsAAAAAAEc0lYadi/BNww9HVpSPqkj1MrmapGqyRr9tM03dTl85IwSqmKomJZbF6uxdpu25yqpmJjvj1wrcvQ33XexPReTO9J+L0N943sTUPk7/SWSB1eA2dj6TXPF+sj/APNP2VrXocL1vYlt/MPOoaEtHdVo/ornFV3GGtkrZI3JrTFajUai8O+uqXwHRAc7ei2rdW6pj1upp3pLiOnWJsX64mmeX1RHJ69gcp03aLbhj262+voLpS0nS0DoXsnY5c83ZoqKndU6sDLct03KdzVyNboGn3tAvxfsTlVHz5VZl6HLEe9f7V4snoPxehyxLvX60eCT1SzQOtwGzsb/AF1xbpx/+YVqoeh2xHFXQSzX21a0yRrn6lJFdki7eSand+csqAZrVii1nufi1OKY1peKTTOkzE7nPL1Zcv8A/gADM1QAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAK/dGv8Akrh/v6T6srBTln+jX/JXD/f0n1ZWCnOppD6TA+WGwg3DIaY8G4ZDTT3OV6/h/Nw/QAYmyAAAAAAAAAAAAAAAACM6R+scPfLfNcSYjOkfrHD3y3zXGWxzkNXjXuF3ubqxdZKDvaPzUMww7F1koO9o/NQzDhV7Uu7o3M0d0fQABxZwAAAAAAAAAAAAAAAAAAfLi/8AaetVJ8gzzUKAOL/2nrVSfIM81DaYd+b5PNPxB5LH6v4soAG0eagAAAAAAAAAAAAAcJ6K6mqKiXDmsU8supbU56hiuy/FcB3YAUX/AAdcOwKrmXegfg64dgVXMu9BegAUX/B1w7AquZd6B+Drh2BVcy70F6ABRf8AB1w7AquZd6DMsdvr0vdCq0NSiJUx5qsTvzk5C7gAAACoOn1MtOGIuWGjX/yUIWTbogEy033/AJaej+qISAAAAAAAAAAAAH49yNarnLkiJmpr/wAOWfPL8JUvOIBsQa/8OWftnSc6g/Ddn7Z0nOoBsAYH4atHbOj55vpH4ZtHbOj55vpAzwYP4YtPbOi59vpP38L2ntpRc+30gZoMJLtal3LnRc+30n0l0ti7lxpOeb6QMsGJ+Erd2fSc830n7+Ebf2dS8630gZQMb8IUHZ1NzrfSOn6Hs2m51vpAyQY/T1F2ZT86npP3pyj7Lg5xAPcHj03S9kw84g6apuyIfHQD2B5dMU/9PF46H7r8P9NH4yAegPjXYv6VnjIfuuR/nt8IH0D8RzV3HIvzn6AAAAAAAAAAAH4rWruoi/MfKxRLuxsX+6h9gDyWmp13YIl/uIfK0dIu7SwL/u0PcAYy0FCu7RU3NN9B+Lbbcu7QUq/7lvoMoAYS2m1Lu2yiX/cN9B8rZbQv/wBMo+Zb6DPAGuWxWdf/AKbS82h+tstrb7mijb8XNDYADEZbqVnuGys+LM9PoU9mQaj3FTWt+LVyp/1HqADXVTfcXO7N+LcZ0+h59tqrmz3F/v7Pi3epT/rPgAZDbnf2e4xZihnxb1Up/wBZ6NvmK2+5xxi9v/PKn1zDAGwbiPGLfc46xZ893mX6XHqzFmN2e5x3if57g5fpNUAN23GmPG+5x3iH56lq/S09G480hN9zj2+p3XRL9MZoABI26RNJTPc4+u/96GmX6Yj1bpN0os9zj6u/vUNI76YiLgCWN0qaVk/n3M741ro/siPRulnSq3+eTXfGtdN9jEIeAJszTBpTbu4npH/GtcX2ZHq3TNpRb/8AXbY741rb9jkIIAJ+3TZpRb/9UsbvjWpfslOp9DrpCxPjetxHSYkdbnrbW0roH0lO6LPXdd1WqRXuz/Fplub5W07R0HnX/G6fBW7/AP6QLFAAAAAAAAGPcKykt1DPX19TDS0tOxZJppXo1kbUTNVVV2kQyDyqqeCqp5KaqgingkarZI5GI5rkXeVF2lQCoWlbopr07FccWj5tPFZ6N66uWrp9WtevKi5KxnBkqOXdVU3E6/oM094d0iujtFdG2zYiy2qV7846jLdWJy7q7+oXbTe1WSqRXTh0M1lvdLUXnAEMVpuzUV7rei6mmqeRqf8AdO4Mva8ibpTmphuFmuz4J2VFBcKKZWuaubJIZGL4UcioRX9Uzz16HX9Y11mu5arUapNVlw5Fd9B2ma8490bXzDuvM2e0FrmdQSLknTqpGqMeiLta4jtTqk3FzRdzPKpNM/FWzNj4HXfZN017VUWTpvX8/H1efzjMyf1CBr8N/hLY7bfwzqfwn0pF05qcste1CavLLay1WZsCoAAAAAAAAAAAAAAAAr90a/5K4f7+k+rKwU5Z/o1/yVw/39J9WVgpzqaQ+kwPlhsINwyGmPBuGQ009zlev4fzcP0AGJsgAAAAAAAAAAAAAAAAjOkfrHD3y3zXEmIzpH6xw98t81xlsc5DV417hd7m6sXWSg72j81DMMOxdZKDvaPzUMw4Ve1Lu6NzNHdH0AAcWcAAAAAAAAAAAAAAAAAAHy4v/aetVJ8gzzUKAOL/ANp61UnyDPNQ2mHfm+TzT8QeSx+r+LKABtHmoAAAAAAAAAAAAAEB0t6RtgL7a38DfhHp5JV/jOtajUaj9B2eeq5NwnxxTonrHerzLh9bRaa+4JElRrnS1O6TUZ63lnqUXLPJfABg+yJ/+z//AMl/6Q9kT/8AZ/8A+S/9I5NsJxlxTvv+Hy+qNhOMuKd9/wAPl9UDrPsif/s//wDJf+kPZE//AGf/APkv/SOTbCcZcU77/h8vqjYTjLinff8AD5fVA6z7In/7P/8AyX/pHtQdEF01XU9NsR1GvStj1X4RzyzXLPLWjkGwnGXFO+/4fL6pl2XBeMI7xRSSYVvbGNqI1c51BKiIiOTbX2oFxwABUvTrQXS4adr1TWmzXO6TvpaNdRR0rpcv4NdtyomTU5XKh52fQzpQubUc+z2uzsXbzuFeiuy+LC1+3yKqFt8k29rdP0CttJ0OeJpmotbja10i77ae1vm/W6Vv0Ge3obanU+20gTq79G0sRPPUsGAK41fQ4X2NFdQ48pJ13mVNoVqeM2X7CJYh0P6S7HG6b8D0V7gbtq+1VOciJ8lIjVVeRquUt2AKEslR00sD2Sw1ELtRNBNGscsTuBzHIitXuoeha/TRozt2OrNJVUsUVLiSljVaCtRMlcqbaRSL76N25ku5nmnLUukldNTtkfE6KTbR8bk22ORcnNXlRUVPmA9QAAPax190w9dvwvh24S2yuXLXHRojo50T3ssa+1end203lQ8QBY/RHphs+KaqCwYnoKSz3+T2sKon/Zq1fgnLttd8G7b4FdvdcWjpF3aWBf8AdoUQqYI6iFYpW5tXb2lyVF3lRd5U4SzfQ24/q8U2KpsF9qFnvdnRiLO73VXTuzSOVeFyZK13KiL74DqS0FCu7RUy/wC6b6D5W225d230i/7lvoMsAYS2m1ru22jX/cN9B8rZrOu7aqBe7Ts9BnkfxJjXCGG62OixBia02qpkjSVkVXVsic5iqqapEcu5mipnyKBnrYbGu7Zrcv8AwrPQfC4dw+u7YrWvdpI/QaFNKmjRf5/YZ+e5xesfaaUdGy/z/wAL/wCKw+sBuFwxhtd3D1pXu0UfoPlcKYWXdw1Zl7tDF6pq00m6N1//AMgYU/xiD1z6TSVo5Xcx/hRf+cU/rgZ64Qwku7heyL3aCL1T4dgvBzvdYTsK923ReqYqaR9Hi7mPMLL/AM3g9c+k0h4AXcxzhhf+bQesB6uwLgh3usHYeXu2yH1T4XAGBF3cFYbXu2uD1QmP8CLuY2w2v/NIPWPtMd4IXcxlh1e5c4fWA8l0eYAXdwNhhe7aYPVPhdG+jxd3AWFl/wCUQeoZSY2wYu5i6wL3LlD6x6sxdhR/uMT2V3cr4l/6gNcujTRwu7o/wp/g9P6h8row0bLu4Awt81pgT/pN7BfLLOuUF4t8qr+ZUsX6FM+N7JG6qN7XtXfauaARBdFmjRd3AOGPmtcKf9J8Loo0Zru4Cw381ui9BNABVzom8I4XwvesIuw5YLdaVqW1yTrSU7Y9c1KQ6nPJNvLNfCpzA7X0YfXXBS9/+bAcUAAAAAANdiVVbYa1zVVFSJdtFLcroK0VL/NVE7lfUp/qFRcT/k9X/Iu+g/oGi5oigc0XQNopX+a707lzq0/1T5XQJopX+bVQncu9Yn+qdOAHLl0A6LF9zYa1vcvFZ96fC6ANGS7lrube5eKr7w6oAOUO6H7Ruu5S3hvcu9R658L0PmjzebfW9y7Tfap1oAcjXoesALuS39vcukn2nw7oeMBruVmIm9y5L9qHXwBxx3Q64HXcueJm9y4J9rD4d0OWC13L5ipvcro/tjOzADiy9Dfg9dzEeLW9ysh+2ErokC0lVXUWvSzJS11TTtklVFe5scz2NzyREzyRN5C+hRG4Jlfr2nBeK7/MyAeYAAAACW6H9H8mkS836mfiCe1R2uKlcxIqZkmrWXXc89VuZa2nhOiO6G6T3uP6tO7bI1/6jD6D78psaJw01v8ApqSxoFendDbV+90gyJ8a0MX/AFEPN3Q23P3ukRnz2RF/1ixIAofWUdRbL1d7PU1LaqS23Goo9fbFraSJG9W6rU5rlnlnlmp8mwxemWkHF6f1/W/XONeAAAAAATDAOirG+NMK02JLVWYdgpKmSZkcdTLM2RNblfGuepYqbasVd3fN0ugPSYm5W4Qd3ayoT/ROr9Cuv+xCzpwVNd/nJjqIFVHaB9J7f+8wm/4twnT6YDpfQ7aO8T4HuOJKvEn4MT8JMpGwNo6h0uWta9qtVqmNy/GNy3d87AAAAAAADiXRHaarlosvFpoaGx0lxbX075XOmlcxWK1yJkmScpyn2X+IOJtr8qk9BaLFeCMI4rngnxJh633WWnarInVMSPVjVXNUQ0nUc0W8RLH5MhBXj2X+IOJtr8qk9A9l/iDiba/KpPQbDoz8C4PwrgmyVeHMO2+1zzXJY5JKaJGq5utuXJeTNEKqhVmvZf4g4m2vyqT0HHNMePY9I2JmYhfh6ktFc6JI6l1PKrkqMvcucip7pE2s99ETgLy0Wh7Re+jhc7AtkVyxtVVWmTbXIqN0XuGbDhXSrDbsO2unttG+2RTOhgRUbq3PkRVy3tpE8AHMMJ3654XxJQYgs86wV1DMksTt5VTdaqb7VTNFTfRVQuZ0Oum2r0oY5rrXXYZt1ulp7c6qSpgernu1MkbNSuabnt893eKnaFrZQXrSvhm1XSmZVUVVcI4p4X56l7VXbRcj+geDtHWCcH3GW44aw7SWyrlhWGSWLVapzFVHK3bVdrNqL8wglKwAVAAAAAAAAAAAAAAAAFfujX/JXD/f0n1ZWCnLP9Gv+SuH+/pPqysFOdTSH0mB8sNhBuGQ0x4NwyGmnucr1/D+bh+gAxNkAAAAAAAAAAAAAAAAEZ0j9Y4e+W+a4kxGdI/WOHvlvmuMtjnIavGvcLvc3Vi6yUHe0fmoZhh2LrJQd7R+ahmHCr2pd3RuZo7o+gADizgAAAAAAAAAAAAAAAAAA+XF/wC09aqT5BnmoUAcX/tPWqk+QZ5qG0w783yeafiDyWP1fxZQANo81AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKKXmWGfFOIqmmVFppr3XSQKm4rFqJFRU5Cy+n/SVDhKySWKz1DX4muMStgYxc1o412lqH8GW3qUX3Tst5Fyq5SQR01NHTxJkyNqNb8wHqAAAAAEy0DXF9r014fcx2pjuMdRQTcrViWVqePE3wkNJToYpJK7TPhOGNFVIKierkVPetZBImfjPanzgXJAAAqx0VLWu0xUWqRF/+X4t1P8AxE5acq10VSoml+gVVRM7BH/mJQOX61F/Rs8VD81mH+ij8VD71Tfzk8J+5pwgeesQf0MfiofnS9P/AEEXiIeoA8elabseHxEHSlL2NDzaHsAPDpOk7Fg5tD86So+xKfm0MgAY/SND2HT8030H5+D6DsGm5pvoMkAYb7Va3+6ttGvdhb6D4js9sifq4aOOF350WbF8KZGeAPWkrLvRZfg/EWIKHLc6XutQxPAj8je2/SDpHt2XSWO7s5E97VsiqUXnGKv6yOgDb42xhizGlTaXYjqLbOy2a9rUlPTOhkesiNRdV7ZWr7hNxE3zUAAAAAAAGuxN+T1f8g76D+gEK5xMXhah/P8AxL+T9f8AIP8AoL/Uq50sS8LE+gD1AAAAAAAAAAAAACid0TLEd+TgvVd/mZC9hRS8JlijEScF7r/8zIB4AAAAAOydCB+VWMk4aW3+dUFjiuHQhflbjBOGjoPOqCx4AAAUfxltaRMXp/X1Z9YprTZ42TLSNi9P69q/PNYAAAAAAWg6FZf9idrTgq67/NynUjlnQqr/ALF7cnBW13+alOpgAAAAAAAAYN+utBY7LWXi6VDaeiooXTzyL71jUzXuryb5SXSV0TWO7/c5o8MVKYdtKOVIWxRtdUPbvK97kXJeRuWW5mu6WK6MKWoi0A33pdXI18tMyVU/MWdn25J85QAkrCQ4pxxi/FNNFTYjxHcrpBFJrkcdTOr2sdllmiLuLkqkeLWdFzb8P2nQhgiiw7T0tPQOqWSU7YURNcYsCrq8090q6pFV2+q8pVelglqamKmgjdJNK9GRsam25yrkiJ84Ezbpc0nNajW47v6NRMkRKx+0nhI3iXEN8xLcEuOILrV3OrbGkSTVMivejEVVRua72ar4T+oVNTRxU0Ubo41VjEavtU3kKR9HM1rdM1MjWo1PwNBtInwkoEB6Hz+W3CH9qRfSf0iP5u9D5/LbhD+1IvpP6RCCQAFQAAAAAAAAAAAAAAABX7o1/wAlcP8Af0n1ZWCnLP8ARr/krh/v6T6srBTnU0h9JgfLDYQbhkNMeDcMhpp7nK9fw/m4foAMTZAAAAAAAAAAAAAAAABGdI/WOHvlvmuJMRnSP1jh75b5rjLY5yGrxr3C73N1YuslB3tH5qGYYdi6yUHe0fmoZhwq9qXd0bmaO6PoAA4s4AAAAAAAAAAAAAAAAAAPlxf+09aqT5BnmoUAcX/tPWqk+QZ5qG0w783yeafiDyWP1fxZQANo81AAAAAAAAAAAAAAAAam44mw3bat1JccQWmjqGZK6Kesjje3NM0zaqoqbW2Yq43wWm7i/D6f8yh9Yqvp7pqebTnipZoIpFTpPJXsRcv+yx8JDekaHsOn5pvoAuwuOsEJu4xw6n/M4fWPlce4FTdxphxP+aQ+sUp6RouxKfm0P3pKj7Eg5tALqLpAwEm7jbDSf81g9Y+V0h4ATdxzhhP+bQesUt6TpOxYObQ/elKXsaHm0AulDpAwFPPHBDjfDUssjkYxjLrArnOVckRER22qrvEmKE01PAy5W1zIImuS40mSoxEX8ewvsANZii+W7DWH6y+3eV0NDRx65M9rFeqJmibSJtrumzIB0RSZ6E8U95f9bQPGg03aLaxqf/N1LSu321kMtOqLwLrjUNuzSfo3dHriY/wtqeW7QIvg1WZTtdvdPNYIFXNYY1Xh1KAWqvWnXRnbmuSC/Ldpk3IrbTvqFd/eRNQnzuQ5fjPT7ie8RyUmFbWzD1M7a6cq1bNVKnC2NM42L3Vf3DlKbSZIAPlGvdUTVM881TVVD9cnqJ5Fklmd+c5y7aqfQAAAAADydOzpqOkibJUVcq6mKmgYsk0i8DWNzVfAB6Oc1rVc5Ua1EzVVXaRDvPQo4PnhgrMfXGB0TrhElNa2PTJyUqO1TpcvhHI3L9FiL740uivQbcbtUwXjH9P0nbmOSSKzapHSTruotQqbSN+DRVVffKm2i2SjYyONscbWsY1ERrWpkiIm8gH0AABH8TYJwhiarjrMQ4atV1qIo9bjlqqZsjmszVdSiqm5mqr85IABCF0R6MF/mHh/yJnoIRpdsWhzR/YG1dRgCw1dzqlWO30DKZrX1D0TbVV96xu0rnb3Kqoi9cxNerdhzD9dfbtPrNFQwummfurkibiJvqq5Iib6qiFLcT4gueMMTVWKL1m2pqfaQU+eaUlOi5siTuZ5uXfcqqBpKS308VdUXFaSjgqqlyucylhSOKJF95G3eam5wruqqqZoAAAAAD7o6avuV0pbRaKKSvudY/UU1NHtK5d9VXca1E21cu0iAeM0scMaySyMjY3dc5ckT5zZYdsGKMSsbJhvC92ukLvc1DYUigd3JZFaxfmVSwui/QbYbA2G64qbBiC+pk7+FZnS0ruCKNdpVT89yK7azTU7h11ERERETJEAqXSaFtKVSxHOtFmo896pumapzbHJ+s9J9CGlCFuaUWH6nkhub0X9qJELYgCl910e6R7S1z67A10exPf0L46tF/uxuV37JFnVcMdWtHUK+lqm+6p6mN0Uqf3Hoi/qL8msxDYLHiGiWivtoobnTr/3dVA2RE5UzTaXlQCkAOjdEHo/w/gW52Gowy2rpILnJOyekfULLC3UMRyKzV5ubtqu1qsuRDnIAAAAABr8SdYK/vd/0F+6Bc6GnXhib9BQXEfWCv73f5ql+LYudtpV4YWeagGSAAABxnTRpnjw7VTYZwi2CuvzPa1NRJ7anoOR2Xu5OBibm67gUOl4vxZhzCNu6fxJeKW3QLtM113t5F4GMTNz15GoqnG8SdEa1znRYSwrPUt3Eq7pL0uzupG1HPVO7qFOG1stXcrpJd7xXVFzucv4yrqXap+XA3eY3ga1ERD8AnF00yaUrg5XMv8Ab7Ui+8oLaxcvnmV6mlmx3pCmXVS48vqqv5j4o08DWIaEAbxMbY9Tcx3iH56hq/8ASfbceaQ2+5x7fU7qwr9MZoABJY9JGk2P8Xj65/36Wld9MRF40nWSeeqqXVNTUzyVE8zmtar5JHq9y5NRETbVdpERD7AAAAAABusB4yxNgW73Gvw7HaJfwhDDFM2vhkfqdbV6ordQ9v567ue4THq+6S+wMIr/AMLUffHNAB0xNP2knft2El/4eo+9P1NP+kffteFF/wBzUfeHMgB9VdXW3K73O73FKdtVcaySrkbAipG1XrmqJmqrl858gAAAAAAFnehUX/Y1RJwV9d/mZDqpynoU1/2O0qcFxrk//syHVgAAAAAAAANNjbDtBi3Cdzw3c0ctJcIHQvVvumKvuXJytVEVOVEP596T9EONsA3KaK42moq7e1y6zcaWJz4JG7yqqe4X9F2S91Ns/o6AP5YUkV4vMtPbqSOuuMkaamnp4mvlVua7jWpnlmvAWi6GPofbnb71SY1x3S9KvpXJLb7Y/JZNcTbbLKnvdTuo3dzyVcssltYjUTPJETPdyQ/SZLmFHOjo/lnpv7Gg+slLxkCx9ohwBjq9tvWJ7LJW1zIGwJI2rmj9o1VVEyY5E3XLtlRRrofP5bcIf2pF9J/SI5lhzQPouw9faO92nD0sFfRSpNBItfO/UvTcXJz1RfnQ6aFAAEAAAAAAAAAAAAAAAAV+6Nf8lcP9/SfVlYKcs/0a/wCSuH+/pPqysFOdTSH0mB8sNhBuGQ0x4NwyGmnucr1/D+bh+gAxNkAAAAAAAAAAAAAAAAEZ0j9Y4e+W+a4kxGdI/WOHvlvmuMtjnIavGvcLvc3Vi6yUHe0fmoZhh2LrJQd7R+ahmHCr2pd3RuZo7o+gADizgAAAAAAAAAAAAAAAAAA+XF/7T1qpPkGeahQBxf8AtPWqk+QZ5qG0w783yeafiDyWP1fxZQANo81AAAAAAAAAAAAAAAAU806/y5Yq/wCD/wArGRAl+nX+XLFX/B/5WMiAAAAAAAi64W7+0aT69hfIobH/AB+3f2hS/XsL5ACA9EPt6E8Vd4r5yE+I1pRw/VYr0fXrDtDNDDU19K6GOSZV1DVVU21yRVy7iAUyB1un6HjGEiJ0ziuxU/DrdFLL9L2mUnQ335UzdpAokXgSyOX/AFwONA65VdDvi+Jq9KYtslUu8k1DLD+tHv8AoIpf9E+kuxxulnw7FdYWJm6S01STLzb0Y9e4iKBDgecczHTSwKj454XamWGViskjXgc12StXuoegAw6+nrppEdS3FaVupyVustft8O2ZgA11mpamjr9evLG4hpd+mdVSUWq5NVFtlgNE+lPRTYWJRuwkmCJ5ERr6lYknik5HVLc3r3ZEROU4gF20yUC9dDV0lfRxVlDUw1VNM3VRTQyI9j28KOTaVO4e5SLBGKcRYFr+nML1iRwPfqqi2zKq0tRw5t9479NuS8OabRavRZpDsmkCzvqbfqqWvpsm11vmVNdpnLufGYu3qXptLyKiogTIAAAABXXotMTOqrnasDU0n8DG1LlckRfdZKrYI1/vI96p+iw4ubPGl3diHH+JL8rtWyquMkcC5/8AcwrrUeXdazP+8prAAAAAADzqZo6ankqJnamONqucvIhZ7oc9HmxbD2yG8U+WIrvG18qPTbpIF22QJwLuK7hd8VDhOibD7MV6UrHZ5mJJR073XKtaqZoscKorWqm+iyOjReTMucAAAAAAAABwDowU28HL/wCJqvqmnDDuvRg+4wev/i6n6pDhQAAAAABgYi6w1/e0nmqX2s652ijXhgZ5qFCsQdYbh3tJ5ql87GudkoV4aaPzUAzQDGuldS2y2VVyrZUhpaSF880i7jGMarnL8yIoHK+iL0kT4Vt0WGsPzozENzjV2vJtrRU+eSzZfnKubWJwoq+9yWs1NAyni1uPVLtq5znLm5zl21cqruqq7aqZd3vVZie/3HFNxRyVN0m15GOXPWYtyKJPisyTu5rvngAAAAAlGijAtVpExTJb1llprJQI190qY1yc7P3MDF3nORM1X3reVUAjVpprheax9HYbTcbxURrlIyhp3SpH8dye1b86oSLqcaTNb13YDdNRu/xml1Xi67mW9w/ZbTh+0wWmyW+noKGBupjhhZqWpy8qrvqu2u+bACi94tOILK1z73hi/W2JqZrNPQSa0n+8ait/WYNNPDUwpNTysljduOY7NFL7FGcSojcc4taiIiJiK45InfMgGGAAAAAxp66mhn1iSRdc1KO1LWK5clzRF2k5FPn8JUn58nMv9B2joSfy9xNwfgul+tmLKgUB/CVH+fJzT/QPwnR/0j+af6C/wAoPS1MFVGr4JNW1rlau0qZKm9tnsbrSImWlDF/9szfQ00oAAAAABZroUV/2QwpwXOu/zDzrByboUP5JGJwXSt+vcdZAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAK/dGv8Akrh/v6T6srBTln+jX/JXD/f0n1ZWCnOppD6TA+WGwg3DIaY8G4ZDTT3OV6/h/Nw/QAYmyAAAAAAAAAAAAAAAACM6R+scPfLfNcSYjOkfrHD3y3zXGWxzkNXjXuF3ubqxdZKDvaPzUMww7F1koO9o/NQzDhV7Uu7o3M0d0fQABxZwAAAAAAAAAAAAAAAAAAfLi/8AaetVJ8gzzUKAOL/2nrVSfIM81DaYd+b5PNPxB5LH6v4soAG0eagAAAAAAAAAAAAAAAKe6d2q3TlijNMtUlG5OVOlmJ9KKQ86F0S9P0tpqmdltVdmpZ/nbJMxf1NQ56AAAAAAfiORtZQOcqI1tfSqqrvIk7C+ZQO7wuntdTFGqo90btQqbqORM0Xw5F5cF3iPEOELPfY1RW3Chhqdre1bEcqfMq5AbcAAAAAAAEL0m6NsOY8ocrhB0rdIm5Utzp2ok8K7yZ+/ZwsdtLyLtpU3ENmuuGcRVmHL7E1lwpFRdWzPW541z1ErM/euyXuKiou2heY4j0W2Hop8K2/GEMaJU2ipbDUPT31NM5GKi8OUixuTg9twgV7AAAAADIst0uuHb9S4hsFQlPc6T3Kr7idi+6ikTfY79S5KmSoY4Audo3xfbcc4Spb/AG1FjSTOOop3rm+mmbtPjdyou/voqLuKSQqZ0PGKX4Y0lQ22aRW2vESpTSNVfasqkRdZf3XJnGvCqs4C2YA1+JK78F4duVz7EpJZ/EYrvsNgRjS0rm6K8XOb7pLHWqnd1h4FJ7FGsdmo2qqq7WWq5V31VM1XwmaeNBklDAibmtN+g9gAAAAADsnQiULZcUYruz27dPTUtJEvxlkkf9EfgLHHBug/ROkMWr778Iwp82sN9KneQAAAAAAAAOB9GAn8BhBf/G1H1Jwk7x0YH8Uwiv8A4+f6lTg4AAAAABhX/rFcO9pPNUvhh5c7Bbl4aWLzEKH33rHX97Seape7DS54cti8NHF5iAbE5X0U10db9ENZRRvVkl3qoLeip+a92qkT542PT5zqhwrowZHLYsK03vHXd8ipytp5cvOUDgSbSZIAAAAA86qZtPSy1D/cxMV69xEzLd6BcMNwroutFJJGja6rjSur3Zbbp5URzkX4qalicjEKf3KFtTTJSv8Aczyxwu7jntav6lL9IiIiIiIiJuIgH6AABRrFG1jzF3/8iuH+YeXlKOYr2sfYu/8A5DX/AF7wMAAAAAB2LoQo1fi7F8+9FR0MfjOqF+wsgcD6D2l/7Li65Im1LXQUqLw63CjlT/zTvgAAAUo0j7WlLF/9sSeaw0hvdJW1pUxh/a7/ADGGiAAAAAALMdCh/JPlwXWt+uU60ck6E/8AkpenBdqz6w62AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAV+6Nf8lcP9/SfVlYKcs/0a/wCSuH+/pPqysFOdTSH0mB8sNhBuGQ0x4NwyGmnucr1/D+bh+gAxNkAAAAAAAAAAAAAAAAEZ0j9Y4e+W+a4kxGdI/WOHvlvmuMtjnIavGvcLvc3Vi6yUHe0fmoZhh2LrJQd7R+ahmHCr2pd3RuZo7o+gADizgAAAAAAAAAAAAAAAAAA+XF/7T1qpPkGeahQBxf8AtPWqk+QZ5qG0w783yeafiDyWP1fxZQANo81AAAAAAAAAAAAAAAAVr6LiiWDG2Gbnqfa1dBU0qrwLG9j0TwSO8CnISyHRaWh1Xo4pr3G1VfZbjFUPVEzXWZM4X/N/CNcvxSt4AAAAAALIdCbfErdHU+H5H51FirHwI1V29YkVZYl7ntnNT4hW8l2hXFbcGaS6Osqpdbtd2alvr1Vcmxqrs4ZV+K9VaqruJIq7wFxAAAAAAAACFadaVtXoZxhG9M9RZ6mZO7HGr0/W1CanOeiRvEdo0OX2NXIk9zh/BsDc9t7p11ConcYr3dxqgVRhdq4mP/Oain0fjGo1qNTcRMkP0AAAAAAxrk+eGjfVUj9bqqZUqIHputkjVHtXwtQvbYLjFd7Fb7tAmUVbSx1DPivajk/UpRifJYJEdualc/AXE0Iukdocwasmeq/AdIm3way3L9WQExNPjaidcsGXu3NbqnVVuqIUThV0bm/abgAUCs0iS2ijkRc9VAxf2UMs9bna1sOILzh9zdT+DLjPTMT4NHqsa/OxWr855AAAAAAHcOg+qEWqxlR77JaOf5nMkb/plgirPQvXRtv0r1NukcjWXe1uazb91LA9HIniPkX5i0wAAAAAAAAHBujA/iGEl/rGb6hxwY710YPW3CS/1nMn/wDXecFAAAAAAMO99Za7veTzVL14VXPC9qXhoofMQorees9b3vJ5ql6MILnhOzrw0MH1bQNqcP6MGBVwnhuuRM0gvaRu5EfTzJn4UTwncDnXRIWeS8aHb2lOxX1NAxlwiRN3OB6SOROVWNenzgVSB8xvbJG2Ri5tciKi8KKfQAAAYd5dKy2TzQ7csKJKzusVHJ9BfW11sFytlLcKV+rgqoWTRO4WuajkXwKUVciORWqmaLtKhZLoV8VsumB1wrVS/wD6jh9Ugajl25KVc9ZenIie0XgVnKgHYQAAKO4u2tIGL/8A+QV31zi8RRO6VbLjiO/XSJyOirrxWVMapvsdO9W/qyA8QAAAPGqjqqhjKGgZq62skZS0rPzpZHIxieFUAtF0KttWi0Q01c5upfda2prlT9FZFYxfnZGxfnOrGtwvaKfD+GrZYqT8Rb6SKljXLdRjUai91cszZAAABSvSbtaVsYJ/Wzvq4zQm/wBJ+1pYxin9ar9VGaAAAAAAAsr0Jq/7LJU4LvV+eddOQ9CYv+y6oTgvFX5yHXgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACv3Rr/krh/v6T6srBTln+jX/ACVw/wB/SfVlYKc6mkPpMD5YbCDcMhpjwbhkNNPc5Xr+H83D9ABibIAAAAAAAAAAAAAAAAIzpH6xw98t81xJiM6R+scPfLfNcZbHOQ1eNe4Xe5urF1koO9o/NQzDDsXWSg72j81DMOFXtS7ujczR3R9AAHFnAAAAAAAAAAAAAAAAAAB8uL/2nrVSfIM81CgDi/8AaetVJ8gzzUNph35vk80/EHksfq/iygAbR5qAAAAAAAAAAAAAAAA12JrRSYgw7cbHXt1VLcKaSmlRN3UvarVVOVM80KOrSVttqqqz3Jupr7dO+kqU4XsXLVJyOTJyLwKhfMrd0VGD3W290+PqGL/slYjKS7alNpkibUMy8ip/BqvIwDjwAAAAAfFRDHUQPgmYj43tVrmrvop9gCwHQ76U218FPgjFNXqbvA3UW6rldtV8SJtNVf6ZqbSpuuRM9tczuRQmohjnjRkiLtKjmqiqjmuTbRyKm2ipvKh1rRvp0vWH4orZjOCovlvZk2O4wIi1cTd7XWbSSp+kmTtrbRygWcBH8H40wri6m1/Dl9orhkmb445MpY/jxrk5vzohIAABHMa44wpg2l1/EV6pqNzkzjg1Wrnl5GRtzc75kAkMj2Rxukkc1jGoquc5ckRE3VVSoum7HjMfYtibbZFfh60Oeyjcm5VTLtPn+Kie1ZyK5ffHtpX0rXrH0clpooJ7Lhty5Pgc5Eqa1OCVUXJjP0EVc99d4gTWta1GtRGtRMkRE2kQD9AAAAAAABj3COonpVpKNmuVVU5tNTsTddJIqMYnhchezD9ths1ht9op/wATQ0sVNH8VjEan6kKsdDzhpcS6Uqeslj1Vvw8xKyZVTadUOzbAzup7Z/8AdbwltQAAAq30UOHnWfSRTYhiZlSX6mSOVyJtJVQplt/Gj1OXyanMS4WmHBrMc4DrbK1zIq5uVRb5nbkVQzbYq8i7bV/RcpTmF0q65FUwPpqqCR0NRA9MnRStXJzF5UVFA9AAAAAHtbLpV2C+W3EVvYslXa6ptSyNFyWVqZo+PP8ASYrm/OXdw9d6C/WOivVrnbUUVbC2aGRN9rkz2+Bd5U3lzQo4TrQ1pMqdHlW+33COaswxUyLJJHEmqkoZFX20kbffMVdtzE28/bN280ULcA1+H71acQWqK62S401wopkzZNA9HNXhReBU30XbTfNgAAAAAAcI6MHrThJf61lT/wDrSHBDvnRg9ZMKL/W7/wDLSnAwAAAAADFvHWis+Qf5ql5cGLng+yrw2+D6tpRq79aqz5B/mqXkwOueC7GvDbqf6toG4PiWNksT4pWNfG9qtc1yZoqLuop9gCkGM8NTYJxlccKTI7WaZ2vW97v+9pHqqxrnvq3bYvKw1ha3Tno5ZjzD8ctA6KnxBbtVJb537TX5+6hf+g7JNveVEXeVFqi5KiGrqKGupZaKvpZFiqqWZMpIXpvLyb6Km0qbaAfoAAGZh683bDGIqTEdhlZHcKXNqsfnrdREuWqiky96uSdxURU20MMAW00b6WcJY0hjgirGWy85IktrrHoyZrt/UZ7UreBzc97NEXaJ8UHqqWmq49bqYI5mcD2ooSnRIOl21FakGWWtJVyozLg1OqyAs7p50p27Dlkq8O2GviqsT1kboWRwvR3SKOTJZpFT3KtTba1dtVy2ss1Kx0sEdNTRU8SZMjYjG9xEPylpqeli1qmgjhZwMaiHsAAAA6Z0M+E34hx4/E1TEq2uwKrYVVPay1rm5bXDrbHKq/pPbwEFwdhu8Y2xIzDtgbqZMkdWVitzjool9+7hcu3qWbqrwIiqXIwZhy14SwzRYfs0KxUdJHqW5rm57l23Pcu+5yqqqvCoG4AAAAAUt0pbWlvGKf1p/oREfJDpV2tLuMk/rNP8vCR4AAAAAAsl0JS/7MatOC81X0tOwHHehIX/AGZ13Jeqr/oOxAAAAAAAAADXYkvdrw5Y6q93qrSkt9I1HzzK1zkYmaJnk1FVdtU3ENiRvSdhbZrgO7YW6e6Q/CESR9Ma1rmt5OR2epzTPc4UAiXsgtD/ABzg8jqPux7ILQ/xzg8jqPuzkXsOk/8A9ir/AIL/AOuY9w6Eiht1K6quGlGno6dvupZ7U1jE7qrPkRXZfZBaH+OcHkdR92PZBaH+OcHkdR92V1rdAeAKelmlbp6wzI+NjnJGkcGblRM8v4zvlfQP6GeyC0P8c4PI6j7skWBdJeCMcVtRRYVv0dyqKaNJZWNglZqWquWeb2om6U50BaCk0qYbr7xsp/A/SlZ0trXSGv6v2jXarPXG5e6yyy3iyGgLQamiq+3G6bKPwx07TJT630hrGoycjs89cdnubmSAdlABUAAAAAAAAAAAAAAAAV+6Nf8AJXD/AH9J9WVgpyz/AEa/5K4f7+k+rKwU51NIfSYHyw2EG4ZDTHg3DIaae5yvX8P5uH6ADE2QAAAAAAAAAAAAAAAARnSP1jh75b5riTEZ0j9Y4e+W+a4y2Ochq8a9wu9zdWLrJQd7R+ahmGHYuslB3tH5qGYcKval3dG5mjuj6AAOLOAAAAAAAAAAAAAAAAAAD5cX/tPWqk+QZ5qFAHF/7T1qpPkGeahtMO/N8nmn4g8lj9X8WUADaPNQAAAAAAAAAAAAAAAAw73bKC9WirtN0pmVNFWROhnifuPY5MlT95mACk+PsI3HAOKX4fuLnzUsiLJa61ybVTCm8q7muMzRHJ3F3FNKXQ0hYOs2OMNy2S9ROViqkkE8e1LTSp7mRi7zk8Cpmi5oqlRMbYXvuBr6lmxFGipIq9JV8bVSCtam+n5r032LtpvZpkoGqAAAAAAAB4T0dNNMyZ8LdeYubJW+1e1eRybaeE3VBifGVvYkdFjfEsUabTWOuDpUanAmuarI1gA2tdirGtexY6zHGJJGLutZXLCi8i63qTRwUlPDM+ZkaLNIub5XKrnvXlcu2vzqe4AAAAAAAAAHnM+RuojghfUVEz2xQQRpm+WRy5NY1N9VVUQVM8VPC6aZ6MY3dVfo5VO/dDvosqaKphxxiykdDXqxfwXQSp7akY5MllkTelcm0ie9RVz21XIOhaFsEtwLgantk6skudQ5aq5TN3HzvRM0RfzWoiMTkbnvk2AAAAAV96JXRvOlRNpAw5SulejE/DVJE3N0jGpklQxE3XNRMnJvtRF3W7dggBQuGSOaJssT2vY9M2uRc0VD6Oz6Z9C1TR1NRibANHr0MjllrbLGmS5rtukp03M99Y9/3u3tLxOnnina5Y1XNrla9rkVrmOTda5F20VOBQPUAAAAB62esudiuDrjh661tnrHe7kpJNSknBq2Lmx/95FOiWfT1pBtcWVypbJfY2Jmr3MfSzOy3c1ZqmeBiHNjzqP4vJ8RfoAu/gy87I8H2bEPS3Sv4ToIKzWdXq9b1yNr9TqskzyzyzyTPgNuRDQqueh7Bn9hUX1DCXgAABwvowk/+X8Kr/XDv8tMcBO/9GF+TWFl/rpf8tMcAAAAAAAMa7daqv5B/mqXhwGueBrAvDbKb6ppR669a6v5B/mqXf0fLngLDy8NrpvqmgbwAADn+ljRXYsexNq3PdbL5CzUU9ygYiuy/MkbuSMz3l203lTNToAApLjfCOKcDTubie1uZRo7JlzpUWSkfwZuyzjXkeiciqaWN7JGI+N7XtXbRzVzRS+UjGSMcx7WuY5FRzXJmipwKc1xVoN0e3yWSpgtktiq5FzdPaZel814VjyWNe6rcwKsg7Jd+hzvsLnOseMqSqZ72K40KscndkjdkviEarNCWlClVUbb7FXfpU1yc3P5pI2gQAEvfon0qtXLYXq+Vt0pvteh9w6JNKkrslwjFDyy3SDL9lygQ0HR6DQRpKqnItQ/DduZv65WSzPT5mxon7RK7L0OGqcj8RY1q5WrtrDbKRlOnc1b1eqp3EQDg9VVU9KxH1EzI0Vck1S7arwIm+pPtHuiTF+M3x1FXBPhuyO23VVVFlVTN+CiXbbn+e/LdzRHFiMF6L8DYRmbU2awU/Tzf/3tSqz1GfCkj1VW9xuSchMgNFgjCdiwZYo7Ph+ibTU7V1cjlXVSTPXdfI5dtzl4V7iZIiIb0AAAAAAApfpY2tMGMk/rFn+XhI6STS3taY8ZJ/WEf+WhI2AAAAAAWP6EZf8AZrcU4L5U/RGdjONdCIv+zi6JwX2p8yM7KAAAAAAAAAAAHF+iV01xaNKKKz2aKGrxJWRa5G2TbjpY81RJHpvqqoupbyKq7SZLSPFuKsRYtubrjiO81lzqVVVR08iq1me81vuWpyNREJH0Qt1qLxpsxbU1L3OdFc5aVma7jIXa01E+ZiEn6D/CNoxbpcSO908dVS22hkr208iZsle17GNRyb6Ismqy/RTPaIqD4f0b4+xBQNr7NhC81lI/3E7KV2of8Vypk75jY9RrSnxFvfk5/SBqI1ERERETaRE3j9GRm4T0GWFsRYUwJeaPEdnq7XUTXPXY46hmpVzNaYmacmaKd2MStuNBRVNJTVdZBBNWyrDTMkeiOmejVcrWpvrk1V+YyyoAAAAAAAAAAAAAAAAAACv3Rr/krh/v6T6srBTln+jX/JXD/f0n1ZWCnOppD6TA+WGwg3DIaY8G4ZDTT3OV6/h/Nw/QAYmyAAAAAAAAAAAAAAAACM6R+scPfLfNcSYjOkfrHD3y3zXGWxzkNXjXuF3ubqxdZKDvaPzUMww7F1koO9o/NQzDhV7Uu7o3M0d0fQABxZwAAAAAAAAAAAAAAAAAAfLi/wDaetVJ8gzzUKAOL/2nrVSfIM81DaYd+b5PNPxB5LH6v4soAG0eagAAAAAAAAAAAAAAAAAAGrxRYLNieyz2a/W+GvoZ09vFKm4u85F3WuTecmSpvG0AFVNI+hbE2FHS12G21GJLKmbtZREWupm8GpT8cicLfbci7pzOmqIahHLDIjla5WvbuOYqbqKi7aLyKX3IRj/RZgzGr3VV0tnS9yVMm3GidrNSndcm09OR6OTkAqIDqWKdAWM7Wr5cO3OhxDTJtpDUf9lqU5EXbjd3V1Bza+2q/wCHlcmIsOXi0o3dlnpXOh+aVmqYvhAxgY9NW0dSiLT1UEuf5kiKZAAAAAAAAAAGLUXChp3Iyarha9VyRmqRXKvAibqkkw7gvHGI1b+BMI3N8TtyprGdKQ5cKOkyVyfFRQNKfdrprhebq2z2G3VN2uT9ynpm5qxPznuX2sbf0nKiHa8I9DrPKrKjG2Ile3dWgtOcbF5HTO9u5PiozunbcK4asGFrY224etNLbaVFzVkDMlev5zl3XO5VVVA5hof0J02H6qDEWL5ILnfI1R9PTsTOmoV4W5+7k/TVNr3qJur2YAAAAAAAAAAc60naIMMY2lfckSSz3xW5JcaRqap+W4krF9rKnd29raVDooApxjLRlj3CTnvrbO68UDdyutTHSplwvh/GM5ckcicJDaWsparNIJ45HNXJzUX2zV4FTdT5y/JGsV4CwZipVfiDDVtrplTLX3wo2ZO5I3J6fMoFMgWNunQ6YKncrrZdMRWn81kNakzE+aZr1y+c0NX0N1Sq/wDYtIEzE3kqLSyRf2XsA4gfE/4l/wAVfoOxS9Dpili/wONLRMnwlrkZ9EqmJUdD3jrUObFiDDcmaZe2jmZ6QO26EVz0OYN/sOk+paTE0GjuyVOG8BWHD9ZLFNUW23w0sskWeoc5jEaqpmiLltb6G/AAADhvRhfkrhhf68T/AC05X8tXp8wHd8fYetdDZayhpamhuKVauq9XqHN1qRmXtUVc/bovzHJk6H7SBv3rDCc/6oHLQdTTofce799wz4s/oPpOh8x3v3/DfNzegDlQOrJ0PeON/EOHU/3Ux9J0PeNt/EeHuYm9IHHrp1sqvkX+apdzRyuej3Da8NppfqmnAp+h2xnNBJC7E1gRJGq1VSmm30y4SxWFrc+z4YtVoklbK+hooaZz2pkjlYxGqqd3IDZAAAAAAAAAAAAAAAAAAAAAAAAAACmWl/a0zYyT/wAdF/lYSNHatJOhTG9+0iX3ENnr8OtorlNHLGyqmmbI3Uwxxqio2NU3WKu6aDqA6Su2GEvKaj7oDmgOmdQHSV2wwl5TUfdH51AtJfZ+EfKqj7kDmgOl9QLSV2fhHyqo+5HUC0ldn4R8qqPuQOgdCEv+zq7pwX6o+riOznPNAeCbzgPCFbar5UUE9XU3KSrzo3vdGjXMjaiZua1c82LvHQwAAAAAAAAAAAoF0W+DKzC+l643FYXfg6+PWuppctpXuy11ufCj1VcuBzV3zn2j/F17wNimlxHYKhsVZT5pk9uqZKxfdMem+1fQqZKiKf0ex9g3D2OcPS2LElC2qpXrqmORdTJC/eex261yfuXNFVCpmP8AoUsX22pkmwfX0t9o1VVZDM9KepanAuq9o7u6pM+BCKmFj6L+1uo2pe8HVkdSiZOWjqWvY5eFEciKnczXumrxZ0XtXLTPhwthGKmlVMm1FfUa5qf92xE2/wC98xx6XQbpZjm1p2CLkrs8s2qxzfCjsjf4e6GjSvdZGpU2mjtETv8AvK2sZtJ8WNXu/UBB7npGxhdccUOMbreaitutDUMnp3SLkyJWuRyNaxMka3a20REzP6O4TvdHiXDFtv8Ab3Z0twpmVEe3ttRyIupXlTcXlQ4Zoy6FjCtimir8XVr8RVbFRyUyM1qlavKmeqf86oi77SwVLBBS00dNSwxwQRNRkccbUa1jU2kRETaRE4BA9QAVAAAAAAAAAAAAAAAAFfujX/JXD/f0n1ZWCnLP9Gv+SuH+/pPqysFOdTSH0mB8sNhBuGQ0x4NwyGmnucr1/D+bh+gAxNkAAAAAAAAAAAAAAAAEZ0j9Y4e+W+a4kxGdI/WOHvlvmuMtjnIavGvcLvc3Vi6yUHe0fmoZhh2LrJQd7R+ahmHCr2pd3RuZo7o+gADizgAAAAAAAAAAAAAAAAAA+XF/7T1qpPkGeahQBxf+09aqT5BnmobTDvzfJ5p+IPJY/V/FlAA2jzUAAAAAAAAAAAAAAAAAAAAAAAAAAEcvuBMFX1zn3jCdkrpHbsk1DG5/jZZp4SJ12gXRhUqrobFUUL199SXGojy/uo/U/qOngDi1V0OOD5FV1NiDFVLwI2rien7cSr+s18vQ3UX/AHGOLy1PhaWB/wBDUO8gCv69Dc7Pax7WZf2bF6x6x9DdB/32OroqfB0ULfpRTvYA4nS9DfhVqo6rxNimpXfalRDG1fFiRf1m9t2gXRjSuR89kqLjInvq24Tyov8Ad1ep/UdPAGlw9hPC+HkRLFh202xUTLVUtIyNy91Wpmvzm6AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACv3Rr/AJK4f7+k+rKwU5Z/o1/yVw/39J9WVgpzqaQ+kwPlhsINwyGmPBuGQ009zlev4fzcP0AGJsgAAAAAAAAAAAAAAAAjOkfrHD3y3zXEmIzpH6xw98t81xlsc5DV417hd7m6sXWSg72j81DMMOxdZKDvaPzUMw4Ve1Lu6NzNHdH0AAcWcAAAAAAAAAAAAAAAAAAHy4v/AGnrVSfIM81CgDi/9p61UnyDPNQ2mHfm+TzT8QeSx+r+LKABtHmoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAK/dGv+SuH+/pPqysFOWf6Nf8lcP9/SfVlYKc6mkPpMD5YbCDcMhpjwbhkNNPc5Xr+H83D9ABibIAAAAAAAAAAAAAAAAIzpH6xw98t81xJiM6R+scPfLfNcZbHOQ1eNe4Xe5urF1koO9o/NQzDDsXWSg72j81DMOFXtS7ujczR3R9AAHFnAAAAAAAAAAAAAAAAAAB8uL/2nrVSfIM81CgDi/wDaetVJ8gzzUNph35vk80/EHksfq/iygAbR5qAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACv3Rr/krh/v6T6srBTln+jX/ACVw/wB/SfVlYKc6mkPpMD5YbCDcMhpjwbhkNNPc5Xr+H83D9ABibIAAAAAAAAAAAAAAAAIzpH6xw98t81xJiM6R+scPfLfNcZbHOQ1eNe4Xe5urF1koO9o/NQzDDsXWSg72j81DMOFXtS7ujczR3R9AAHFnAAAAAAAAAAAAAAAAAAB8uL/2nrVSfIM81CgDi/8AaetVJ8gzzUNph35vk80/EHksfq/iygAbR5qAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACv3Rr/krh/v6T6srBTln+jX/JXD/f0n1ZWCnOppD6TA+WGwg3DIaY8G4ZDTT3OV6/h/Nw/QAYmyAAAAAAAAAAAAAAAACM6R+scPfLfNcSYjOkfrHD3y3zXGWxzkNXjXuF3ubqxdZKDvaPzUMww7F1koO9o/NQzDhV7Uu7o3M0d0fQABxZwAAAAAAAAAAAAAAAAAAfLi/wDaetVJ8gzzUKAOLu6I8QwYm0e2m4xSI6VkDYKlue22ViI1yLwZ5ZpyKhs8OmM5h5x+IFqqbVm5EeqJmPHLL6JYADavMQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAV+6Nb8lsP9/SfVlYKc7z0ZuJaetxBaMMU0iPfb431FVkuepfJqdS1eVGtz7j0OD06HT0iX02B0znDPh3DIaY8O4ZDTT3OV69h8f+OH6ADG2IAAAAAAAAAAAAAAAARnSP1jh75b5riTEZ0j9Y4e+W+a4y2Ochq8a9wu9zdWLrJQd7R+ahmGHYuslB3tH5qGYcKval3dG5mjuj6AAOLOAAAAAAAAAAAAAAAAAAD8cS3RfpDvGALs+oompVUM+XTVHI7JsmW4qL71ycPhRSJqfD0zMluuaJzh0NP0S3pVmbVyM4lbmw6d9Hdypmvq7nPap1T20NVTvXJeRzEc1U+dO4bNdMmjNN3FtJzcnqlK5Y8zHfFyG0o0yZj1vMtN9FKLdU73M5f72LudWXRlxto+bk9UdWXRlxto+bk9Uo+sPINZ5DLwlrJ9Hqo2rwdWXRlxto+bk9UdWXRlxto+bk9Uo/rPINZ5BwlNXq+1eDqy6MuNtHzcnqjqy6MuNtHzcnqlH9Z5BrPIOEmr1favB1ZdGXG2j5uT1R1ZdGXG2j5uT1Sj+s8g1nkHCTV6vtXg6sujLjbR83J6o6sujLjbR83J6pR/WeQazyDhJq9X2rwdWXRlxto+bk9UdWXRlxto+bk9Uo/rPINZ5Bwk1er7V4OrLoy420fNyeqOrLoy420fNyeqUf1nkGs8g4SavV9q8HVl0ZcbaPm5PVHVl0ZcbaPm5PVKP6zyDWeQcJNXq+1eDqy6MuNtHzcnqjqy6MuNtHzcnqlH9Z5BrPIOEmr1favB1ZdGXG2j5uT1R1ZdGXG2j5uT1Sj+s8g1nkHCTV6vtXg6sujLjbR83J6o6sujLjbR83J6pR/WeQazyDhJq9X2rwdWXRlxto+bk9UdWXRlxto+bk9Uo/rPINZ5Bwk1er7V4OrLoy420fNyeqOrLoy420fNyeqUf1nkGs8g4SavV9q8HVl0ZcbaPm5PVHVl0ZcbaPm5PVKP6zyDWeQcJNXq+1eDqy6MuNtHzcnqjqy6MuNtHzcnqlH9Z5BrPIOEmr1favB1ZdGXG2j5uT1R1ZdGXG2j5uT1Sj+s8g1nkHCTV6vtXg6sujLjbR83J6o6sujLjbR83J6pR/WeQazyDhJq9X2rwdWXRlxto+bk9UdWXRlxto+bk9Uo/rPINZ5Bwk1er7V4OrLoy420fNyeqOrLoy420fNyeqUf1nkGs8g4SavV9q8HVl0ZcbaPm5PVHVl0ZcbaPm5PVKP6zyDWeQcJNXq+1eDqy6MuNtHzcnqjqy6MuNtHzcnqlH9Z5BrPIOEmr1favB1ZdGXG2j5uT1R1ZdGXG2j5uT1Sj+s8g1nkHCTV6vtXg6sujLjbR83J6o6sujLjbR83J6pR/WeQazyDhJq9X2rwdWXRlxto+bk9UdWXRlxto+bk9Uo/rPINZ5Bwk1er7V4OrLoy420fNyeqOrLoy420fNyeqUf1nkGs8g4SavV9q8HVl0ZcbaPm5PVHVl0ZcbaPm5PVKP6zyDWeQcJNXq+1eDqy6MuNtHzcnqjqy6MuNtHzcnqlH9Z5BrPIOEmr1favB1ZdGXG2j5uT1R1ZdGXG2j5uT1Sj+s8g1nkHCTV6vtXg6sujLjbR83J6o6sujLjbR83J6pR/WeQazyDhJq9X2rwdWXRlxto+bk9UdWXRlxto+bk9Uo/rPINZ5Bwk1er7V4OrLoy420fNyeqOrLoy420fNyeqUf1nkGs8g4SavV9q8HVl0ZcbaPm5PVHVl0ZcbaPm5PVKP6zyDWeQcJNXq+1eDqy6MuNtHzcnqjqy6MuNtHzcnqlH9Z5BrPIOEmr1favB1ZdGXG2j5uT1R1ZdGXG2j5uT1Sj+s8g1nkHCTV6vtXg6sujLjbR83J6o6sujLjbR83J6pR/WeQazyDhJq9X2rwdWXRlxto+bk9UdWXRlxto+bk9Uo/rPINZ5Bwk1er7V4OrLoy420fNyeqOrLoy420fNyeqUf1nkGs8g4SavV9q8HVl0ZcbaPm5PVHVl0ZcbaPm5PVKP6zyDWeQcJNXq+1eDqy6MuNtHzcnqjqy6MuNtHzcnqlH9Z5BrPIOEmr1favB1ZdGXG2j5uT1R1ZdGXG2j5uT1Sj+s8g1nkHCTV6vtXg6sujLjbR83J6o6sujLjbR83J6pR/WeQazyDhJq9X2rwdWXRlxto+bk9UdWXRlxto+bk9Uo/rPINZ5Bwk1er7V4OrLoy420fNyeqOrLoy420fNyeqUf1nkGs8g4SavV9q8HVl0ZcbaPm5PVHVl0ZcbaPm5PVKP6zyDWeQcJNXq+1eDqy6MuNtHzcnqjqy6MuNtHzcnqlH9Z5BrPIOEmr1favB1ZdGXG2j5uT1R1ZdGXG2j5uT1Sj+s8g1nkHCTV6vtXg6sujLjbR83J6o6sujLjbR83J6pR/WeQazyDhJq9X2rwdWXRlxto+bk9UdWXRlxto+bk9Uo/rPINZ5Bwk1er7V4OrLoy420fNyeqOrLoy420fNyeqUf1nkGs8g4SavV9q8HVl0ZcbaPm5PVHVl0ZcbaPm5PVKP6zyDWeQcJNXq+1eDqy6MuNtHzcnqjqy6MuNtHzcnqlH9Z5BrPIOEmr1favB1ZdGXG2j5uT1R1ZdGXG2j5uT1Sj+s8g1nkHCTV6vtXg6sujLjbR83J6o6sujLjbR83J6pR/WeQazyDhJq9X2rwdWXRlxto+bk9UdWXRlxto+bk9Uo/rPINZ5Bwk1er7V4OrLoy420fNyeqOrLoy420fNyeqUf1nkGs8g4SavV9q8HVl0ZcbaPm5PVHVl0ZcbaPm5PVKP6zyDWeQcJNXq+1eDqy6MuNtHzcnqjqy6MuNtHzcnqlH9Z5BrPIOEmr1fau8/TPoxa1XLi2lyTgilVf1NOfaR+iQtFNRS0eCKeWurHorW1tRErIYv0kYvtnLyKiJ3dwrFrPIfqQ8hJ0lzo9H6on15vyuqqy5XCevr6iSpqqiRZJpZFzc9yrmqqp6QMDIuQyI48jq3bub6fC8JqoqiZh6RJkh7IfLUPs19U5vvdHt7inIABxdkAAAAAAAAAAAAAAAAIzpH6xw98t81xJiM6R+scPfLfNcZbHOQ1eNe4Xe5urF1koO9o/NQzDDsXWSg72j81DMOFXtS7ujczR3R9AAHFnAAAAAAAAAAAAAAAAAAAPxUP0BJjN8K3M+FYh7DI5RVkwV6PTVysfW04BracB75DI5buWGdCo2PDW04BracB75DIbuU4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgGtpwHvkMhu5OA29jw1tOAa2nAe+QyG7k4Db2PDW04BracB75DIbuTgNvY8NbTgP3W0PbIZDdysaFRseSMPtrcj7yBxmrNmosU0cgiAA4s8QAAKAAAAAAAAAAAAAAAAEZ0j9Y4e+W+a4kxGdI/WOHvlvmuMtjnIavGvcLvc3Vi6yUHe0fmoZhh2LrJQd7R+ahmHCr2pd3RuZo7o+gADizgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEZ0j9Y4e+W+a4kxGdI/WOHvlvmuMtjnIavGvcLvc3Vi6yUHe0fmoZhh2LrJQd7R+ahmHCr2pd3RuZo7o+gADizgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEZ0j9Y4e+W+a4kxGdI/WOHvlvmuMtjnIavGvcLvc3Vi6yUHe0fmoZhh2LrJQd7R+ahmHCr2pd3RuZo7o+gADizgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEZ0j9Y4e+W+a4kxGdI/WOHvlvmuMtjnIavGvcLvc3Vi6yUHe0fmoZhh2LrJQd7R+ahmHCr2pd3RuZo7o+gADizgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEZ0j9Y4e+W+a4kxGdI/WOHvlvmuMtjnIavGvcLvc3Vi6yUHe0fmoZhh2LrJQd7R+ahmHCr2pd3RuZo7o+gADizgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEZ0j9Y4e+W+a4kxGdI/WOHvlvmuMtjnIavGvcLvc3Vi6yUHe0fmoZhwyHF+I4YWQxXSRscbUa1NQzaREyRNw+9mmJ+20nNs9B3Jw+5M55w+VtenmgUW6aZt1+qI+EeZ3AHD9mmJ+20nNs9A2aYn7bSc2z0E4uubYZNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv+H9XX4R5ncAcP2aYn7bSc2z0DZpifttJzbPQOLrm2DX/AA/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv8Ah/V1+EeZ3AHD9mmJ+20nNs9A2aYn7bSc2z0Di65tg1/w/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv+H9XX4R5ncAcP2aYn7bSc2z0DZpifttJzbPQOLrm2DX/AA/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv8Ah/V1+EeZ3AHD9mmJ+20nNs9A2aYn7bSc2z0Di65tg1/w/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv+H9XX4R5ncAcP2aYn7bSc2z0DZpifttJzbPQOLrm2DX/AA/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv8Ah/V1+EeZ3AHD9mmJ+20nNs9A2aYn7bSc2z0Di65tg1/w/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv+H9XX4R5ncAcP2aYn7bSc2z0DZpifttJzbPQOLrm2DX/AA/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv8Ah/V1+EeZ3AHD9mmJ+20nNs9A2aYn7bSc2z0Di65tg1/w/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv+H9XX4R5ncAcP2aYn7bSc2z0DZpifttJzbPQOLrm2DX/AA/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv8Ah/V1+EeZ3AHD9mmJ+20nNs9A2aYn7bSc2z0Di65tg1/w/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv+H9XX4R5ncAcP2aYn7bSc2z0DZpifttJzbPQOLrm2DX/AA/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv8Ah/V1+EeZ3AHD9mmJ+20nNs9A2aYn7bSc2z0Di65tg1/w/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv+H9XX4R5ncAcP2aYn7bSc2z0DZpifttJzbPQOLrm2DX/AA/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv8Ah/V1+EeZ3AHD9mmJ+20nNs9A2aYn7bSc2z0Di65tg1/w/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv+H9XX4R5ncAcP2aYn7bSc2z0DZpifttJzbPQOLrm2DX/AA/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv8Ah/V1+EeZ3AHD9mmJ+20nNs9A2aYn7bSc2z0Di65tg1/w/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv+H9XX4R5ncAcP2aYn7bSc2z0DZpifttJzbPQOLrm2DX/AA/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv8Ah/V1+EeZ3AHD9mmJ+20nNs9A2aYn7bSc2z0Di65tg1/w/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4A4fs0xP22k5tnoGzTE/baTm2egcXXNsGv+H9XX4R5ncAcP2aYn7bSc2z0DZpifttJzbPQOLrm2DX/AA/q6/CPM7gDh+zTE/baTm2egbNMT9tpObZ6Bxdc2wa/4f1dfhHmdwBw/ZpifttJzbPQNmmJ+20nNs9A4uubYNf8P6uvwjzO4EZ0j9Y4e+W+a45rs0xP22k5tnoMe4YmvtfCkNZcHyxtdqkRWNTbyVM9pOVTnb0C5TVEzMOniHpvoWk6NXapoqzmNkfdpwAbV5gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA9KePXaiOL896N8KkkrsPU8jVdSOWJ/5qrm1ftQCLg9qummpZVinjVjk8C9w8QAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD7hhlmfqIY3SO4GpmZqWW5q3PpZfHb6QNeDJqKGsp0VZaaRqJv6nNPCYwAAAAAAAAAAAAAAAAAAAAfrGue5GtarnLuIiZqpnR2e5PbqkpXIn6TkRf1qBgAzJrZXwpm+lky4WpqvoMNdpclAAAAAAAAAAAAAAAAAAErbYaKSmj1TZI5NQmqVrt1cuUCKA3lXh2oYiuppWyp+avtV9BppopIZFjlY5j03UVMlA+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACQ4ftEUkLauqbq0d7hi7mXCoEeBPulqfUajWItTwahMiPYitMdPH03TN1LM/bs3k5UJmZNEACgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB+tRXORrUVVVckRN8/CT4YtzY4UrZW5yP8AcZ+9Th+cDEoMPTSNR9VJrSL7xEzd+42cdgtzUycyR/K56/YbQ85qiCD8dNHHn+c5EIrXS2C3vTJrZI14Wv8ASaq4WCohar6d2vtT3uWTv3klgqaef8TPHJl+a5FPUDnqoqKqKmSofhLL9aW1TFngajahEzVE9/8AvIoqKiqipkqbqFR+AAAAAAAAz7BHrl3p03kdqvAmZNCC2+rkoqjX4msc7LLJybRu6bEkaqiVFO5vKxc/1KRYbqrpoKqLW540e3ez3U7hG7rYpYEWWlVZY03W++T0kho66lq0/gJmuX83cXwGQBzwEpvlmbUI6opWo2bdc1Nx/wC8i6orVVFRUVNpUUqPwJtrkgJPhegjbTpWyNR0j1XUZ+9Th7oGibb65zNUlJMqfEUxnNcxytc1WuTdRUyVDoRrMQUEdVRvla1Emjbqkcm+ibxM1yQ8AFQAAAAAAAAAAAAAAAANpZrRLWqksuccHDvu7npPrD1tSslWaZP4CNdz85eAlrURrUa1ERETJETeIrypaaCliSOCNrG8m/3T1PieaKCNZJpGxtTfcuRgPvttauSTOdyoxQNkYNdaaKrRVdEkb19+zaX958xXm2yLklSjV/SaqGdFJHK3VxPa9vC1c0AiNys1VSZvamvRJ75qbad1DWHQzVXOyU9VnJFlDNwontV7qDMyREGRW0dRRy63PGrV3l3l7imOVAAAAAAAAAAADMtlvnr5dTGmpYnunruJ+8/LTRPrqtIkVUYm293AhNKaCKnhbDCxGMbuIhFeNvt9NQx5QszdvvX3SmUfjnNa1XOcjWptqqrtIa+W9W6N2p1/Vqn5rVX9YGxMasoKSrT+Hha5fzk2nJ85jsvltcuWvq3usUzYKiCdM4ZmSJ+i7PICNXGwTw5vpXLMz833yek0zkVqq1yKipuop0IwrlbKWuaqyN1Mm9I3d/eMzJCQZ1ztlTQuze3Vx57Ujdz5+AwSoAAAAAAAAAAD1o49eq4Yvz3o3wqT4gNLM6nqGTsRquYuaI5No3tPiXcSopvnjd9i+kkrCQmFdrdFXw6l2TZE9w/g/cfVFc6KrySKZNWvvHbS/vMsCAVMElPO6GZupe1clQ8yZ3q2sr4c0ybO1PaO4eReQh80b4ZXRStVr2rkqLvFR8AAAAAAAAAAAAAAAAAAAAAAAAAAAAABPLarVt9OrPc603LwEDNzYrwlIzpeoRXQ5+1cm639xJISow72rW2mpV+5qFT5979Z8reLbqNV003LuLn4MiP327LXKkMKK2Bq57e65QrVAAqAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA+o2q+RrE3XKiHQGNaxjWNTJrUyROQgED9RMx6+9ci/rOgIqKmabaKSVhr7/WuoqHVR7Uj11LV4OUhz3uker3uVzl21VVzVST4uic6hjkTbRj9v50IsIJfrHOY5HMcrXJuKi5KhJbDeVnc2lq3Jri7TH/nci8pGT9RVRc0XJUKjoRGcVUCRyJWxNya9cpETedw/ObWwV/TtJk9f4aPafy8CmZVwMqaaSB/uXty7nKRUBB9zRuilfE9MnMcrV7qHwVAAAAAAAAH61Va5HNVUVNxUN5ab9JG5Iq1VezcST3yd3hNEAOhMc17EexyOa5M0VF2lNBii2oqLXQt20/Gon0mHh+6LSSpTzO/gHLtKvvF4e4StyNexWuRHNcmSpwoRXPSZ4dlbLaYUau2zNrk4FzIvd6NaKufDt6hfbMXhQW2vnoJVfEqK1fdMXcUqJweFxmZBQzSvXaRi/Ou8hqW4lg1GbqaRHcCOTLwmout0nuDkRyJHE1c0Yi/rXhJkubAABUAAAAAAAAAAAAAAAATm0wJTW6CJEyXUoru6u2pkSPbHG6R65Naiqq8iH6xUVjVTcVNoxL4qpaanLd1BFRK5VstdUrLIq6n3jd5qGKAVA9IJpYH6uGR0buFq5HmAJBbcQuRUjrm5p/SNTb+dPQSGKRksaSRvR7Hbiou0pz4zbXcZ6CXNi6qNV9sxV2l9CkyXNMqmCKoiWKZiPYu8pFLzaZKFVljzkgVd3fbyL6SU0VVDWQJNC7Nq7qb6LwKer2texWPajmqmSou4oHPQbK/W5aGozjzWCTbavByGtKgAAAAAAACWYUgSO3LNl7aVyrnyJtek25gYeVFs9PlwL9KmeRUSxFcH1NW6BjlSGNcsk98qbqmpP16qrlV26q7Z+FQP1jnMcjmOVrk3FRclQ/ABubdf6iFUZUpr8fD75PSSSjqoKuLXYJEe3f4U7pAj2o6makmSWB6tcm7wKnApMhPHta9qte1HNVMlRUzRSN3qxrGjqiiRVZuuj307huLTcYq+HNPayt92zg5U5DNCueA32JrYkedbTtyaq/wjU3l4TQlQAAAAAAAAAAA21rvdRTKkdQqzQ8vum9xTUgDoFPNFUQtlhej2O3FQ1eI7alVAtRE3+GjTby98nAaKzXF9BUbeboXL7dv2pykzjeySNsjHI5rkzRU30IrnoNpiOi6UrdWxuUUvtm8i76GrKgAAAAAAAAAAAAAAAACUUdjoZaSGV+u6p8bXLk7fVD12P2/4XxxmZIkCW7H7f8L442P2/wCF8cma5IkCW7H7f8L442P2/wCF8cZmSJAlux+3/C+ONj9v+F8cZmSJAlux+3/C+ONj9v8AhfHGZkiQJbsft/wvjjY/b/hfHGZkiQJbsft/wvjjY/b/AIXxxmZIkCW7H7f8L442P2/4XxxmZIkCW7H7f8L442P2/wCF8cZmSJAlux+3/C+ONj9v+F8cZmSJAlux+3/C+ONj9v8AhfHGZkiQJbsft/wvjjY/b/hfHGZkiQJbsft/wvjjY/b/AIXxxmZIkCW7H7f8L442P2/4XxxmZIkCW7H7f8L442P2/wCF8cZmSJAlux+3/C+ONj9v+F8cZmSJAlux+3/C+ONj9v8AhfHGZkiQJbsft/wvjjY/b/hfHGZkiQJbsft/wvjjY/b/AIXxxmZIkCW7H7f8L442P2/4XxxmZIkCW7H7f8L442P2/wCF8cZmSJAlux+3/C+ONj9v+F8cZmSJAlux+3/C+ONj9v8AhfHGZkiQJbsft/wvjjY/b/hfHGZkiQJbsft/wvjjY/b/AIXxxmZIkCW7H7f8L442P2/4XxxmZIkD6kRGyOam4iqh8lQAPahjbNWwRPz1L5GtXLgVQPEEt2P2/wCF8cbH7f8AC+OTNckSBLdj9v8AhfHGx+3/AAvjjMyRIEt2P2/4Xxxsft/wvjjMyRIEt2P2/wCF8cbH7f8AC+OMzJEiY4drEqre1qu/hIk1Dk5N5Tz2P2/4Xxxsft/wvjgbKphjqIHwyJm16ZKQeupZKOpfBKm23cXhThJNsft/wvjjY/b/AIXxwIkCW7H7f8L442P2/wCF8cZmTQWSrWkuEb1XJjl1L+4pNTU7H7f8L45tWNRjGsRVVGplmu6BFMVU+tXFJUT2szc/nTaX7DUE5uFBT1zWNnR3tFVUVq5GHsft/wAL44zMkSBLdj9v+F8cbH7f8L44zMkSBLdj9v8AhfHGx+3/AAvjjMyRIEt2P2/4Xxxsft/wvjjMyRIEt2P2/wCF8cbH7f8AC+OMzJEiT4XuGux9JzO9uxPaKu+3g+Y9tj9v+F8c9KeyUcEzJolla9i5ouqA88UUmv0GvtT28O3/AHd/0kSOhOajmq1yZoqZKhqtj9v4JfHAiQJbsft/wvjjY/b/AIXxxmZIkCW7H7f8L442P2/4XxxmZIkCW7H7f8L442P2/wCF8cZmSJAlux+3/C+ONj9v+F8cZmSJAlux+3/C+ONj9v8AhfHGZkiQJbsft/wvjmqxFbqehZCsGr9uq56pc9zIuaNOAAABsbBRw1ta6KbValI1d7VctvNPSBJLFUpU2yJ2ebmJqHd1DLqYmz08kLtx7VavzmPb7fBQq/WFkyfuo52aGWRXP54nwTPhkTJ7FyVD4JtX2qjrJUlma5H5ZKrVyz7pjbH7f8L44zMkSBLdj9v+F8cbH7f8L44zMkSBLdj9v+F8cbH7f8L44zMketVdJQVKSNzVi7T28KekmcM0U0DZ43osbkzReQ1ux+3/AAvjns20UzaZ1M2SdInLmrdXtKoGoxFdYqpvSsDUcxHZrIu+vJ6TRkt2P2/4Xxxsft/wvjgRIEt2P2/4Xxxsft/wvjjMyRIEt2P2/wCF8cbH7f8AC+OMzJEgS3Y/b/hfHGx+3/C+OMzJ44RqUfSyUyr7aN2qRORf3/Sbw19HaKWknSaFZUcm1tu2lQ2AEJvdK6luMrMsmOXVs7imETq4UNPXMa2dqrqVzRUXJUMLY/b/AIXxxmZIkCW7H7f8L442P2/4XxxmZIkCW7H7f8L442P2/wCF8cZmSL008tNO2aF2pe1dpSaWytjrqVszNp249v5qmHsft/wvjntT2imp9XrEk8erbqXZP3UAxb/dooo5KSFGyyORWv30b+8ixLdj9v8AhfHGx+3/AAvjgRIEt2P2/wCF8cbH7f8AC+OMzJEgS3Y/b/hfHGx+3/C+OMzJEgS3Y/b/AIXxxsft/wAL44zMkSBLdj9v+F8cbH7f8L44zMkSBLdj9v8AhfHGx+3/AAvjjMyRI32F7hrcnSUrvaOX+DVd5eD5zP2P2/4Xxz9bYKBrkc3XkVFzRUeBkXuk6ct8kaJm9vtmd1CEnQ03N3M1ktioJJHSK2RFcqqqI7aAh4Jbsft/wvjjY/b/AIXxxmZIkCW7H7f8L442P2/4XxxmZIkCW7H7f8L442P2/wCF8cZmSJAlux+3/C+ONj9v+F8cZmSJAlux+3/C+ONj9v8AhfHGZkiQJbsft/wvjmBfbVS0dIyWHV6pZEaubs9rJfQXNHpSYgghpYonU8iqxiNVUVN5Mj12SU3Y8vhQjAGQk+ySm7Hl8KDZJTdjy+FCMAZGaY2y7xV9QsLInsVGq7N2XCnpNkRXCPXN/wAkv0oSoitfdLrFb5WRyRPerm5+1yMPZJTdjy+FDFxh/G4Pk/tNEBJ9klN2PL4UMm23mGuqUgZDI1clXNVQh5tsK9dk+I4CWmBdbnFb3Rtkje/VoqpqctrIzyOYy/GU3cd9gHvskpux5fCg2SU/Y0vhQjALkmaT7JKbseXwoNklN2PL4UIwBkZpPskpux5fChtqOpiq6ds8Ls2u8KLwKQI2mHa5aStSN7v4GVdS7kXeUmS5peaepv0VPO+GSmlR7FyXbQ3BHsXUifwdYxP0H/Yv/vkA9dklN2PL4UGySm7Hl8KEYBckzT2iqG1VKyoa1Wo9M0Rd3dPYwcP9Z6f4q/SpnHFWkfiOna9zVp5dpct1D82SU3Y8vhQjc/4+T4y/SfByyTNJ9klN2PL4UGySm7Hl8KEYAyM0n2SU3Y8vhQbJKbseXwoRgDIzSfZJTdjy+FBskpux5fChGAMjNJ9klN2PL4UGySm7Hl8KEYAyM0n2SU3Y8vhQzLVdYrhI9kcT2K1M11WRDDe4O/jU/wARPpIqTGtud3ioKhIXxPeqtR2bcv8A3vGyIpi3rm35JPpUDP2SU3Y8vhQbJKbseXwoRgFyTNJ9klN2PL4UGySm7Hl8KEYAyM0n2SU3Y8vhQ2Nrr47hE+SNjmI12pycQclGD/4lN8p9iEVuzU1l9gpqqSB0EjlYuSqiobYhV9671PxvsBLc7JKbseXwoNklN2PL4UIwC5Jmm1quMdwZI6ONzNQqIuqM00ODvxNT8Zv0Kb4itVX3uGkq3074ZHOZlmqKmW2mf2nhskp+x5fChqcR9eqj+75qGuA+pHap7ncKqp8gFQPWklSCqimVNUkb0dlw5KeQAkyYlgy26aRF5FQ/dklN2PL4UIwBkZpPskpux5fCh7Ul/o5pUY9Hw57jnZZfuIkBkZuhoqKmaLmihc8lyTNSJ2G7PpZGwTuVady5Iq+8/cSwitNNiCKGV0UtJM17VyVFVD42SU3Y8vhQ+sUUKTU3TcafwkSe25W/uIsBJ9klN2PL4UGySm7Hl8KEYBckzdCauqajuFMzzq5201NJO5quRiZqiH3F+KZ8VDFvfWmp+IpFa/ZJTdjy+FBskpux5fChGAXJM0oZiKne9rUp5dtct1DdkAp/x8fxk+kn5JWHhcKptHSPqHtVzW5bSbu2uRqtklN2PL4UMvEvWabut85CGglJ9klN2PL4UGySm7Hl8KEYBckzSfZJTdjy+FBskpux5fChGAMjNLaK+wVVVHTtgkar1yRVVDbEKsPXen+N9ik1JKsO6V8dviZJJG56Odl7U1+ySm7Hl8KDGH8Ug+U+wjAEn2SU3Y8vhQ9aW/QVFTHA2CRFe5Goqqm0RMy7N11pvlELkJwYt0rmUFO2aRjno52pyb3FX7DKNNi/rbH8snmuIPjZJTdjy+FD82SU+f8AFpcu6hGQXJM0n2SU3Y8vhQbJKbseXwoRgDIzSfZJTdjy+FBskpux5fChGAMjNJ9klN2PL4UNla65lfA6aNjmI12pyd3EX7SDEqwh1ul+WX6EIrcmkdiOna5W9Ly7S5bqG7OfzfjX/GUEpJskpux5fChrL7c47gkSMiczUKq+2XdzyNWCoAAAZ1lrWUFW6Z7HPRWK3JO6noMEASfZJTdjy+FDbUVQ2qpWVDWq1HpmiLukCJrYOs9P8VfpUkqzjR7JKbseXwobw54IJSfZJTdjy+FBskpux5fChGAXJM0n2SU3Y8vhQbJKbseXwoRgDIzSfZJTdjy+FBskpux5fChGAMjNJ9klN2PL4UGySm7Hl8KEYAyM0n2SU3Y8vhQbJKbseXwoRgDIzTW1XOO4OkSON7NRlnqt/MziOYN/GVPcb9pIyK1lxvMNFUrA+GRyoiLmipvmNskpux5fChrcU9dnfEaaoCT7JKbseXwobC118dwifJGxzEa7L2xCCT4P/ic3yn2Abw1NbfYKWqkgdBI5WLkqoqG2IVfuu9R8b7AS3OySm7Hl8KDZJTdjy+FCMAuSZpPskpux5fCg2SU3Y8vhQjAGRmk+ySm7Hl8KG1t9S2spGVDGq1rs9pd3aXIgZMsNdZoP73nKSVbE0j8R07XuatPLtLluobs5/P8Aj5PjL9IglJNklN2PL4UGySm7Hl8KEYBckzTu3VbK2lSdjHNRVVMl5DINXhbrQz47vpNoRWmmxDTxTPjWnlVWOVqrmm8fGySm7Hl8KEerv49P8o76TxLkmaT7JKbseXwoNklN2PL4UIwBkZpPskpux5fCg2SU3Y8vhQjAGRmk+ySm7Hl8KGbarnFcHSJHG9moRFXVZbeZCyQYN/GVPcb9pFSM1FXfoKapkgdBI5WLkqoqG3IRe+u1T8dQS3eySm7Hl8KDZJT5/wAWl8KEYBckzSfZJTdjy+FBskpux5fChGAMjNJ9klN2PL4UGySm7Hl8KEdhpqib8TBJJ8VqqZKWm49iP/UBudklN2PL4UGySm7Hl8KGm/BNx7Ef+ofgm49iP/URUyp5Ump45kRUSRiORF3s0zPiuqG0lK+oe1XNZlmibu7kLex0dBTxvTJzYmoqcCoiGNiHrNUdxPOQDC2SU3Y8vhQwrzeIq6mbEyF7VR6OzcqcC+k0oKgAfrUVzka1FVV3EQD8BsqayXCZEVYkiRd+Rcv1bpmMw1Oqe2qY0XkaqgeWEeub/kl+lCVGos1nfQVTpnTtkRWK3JG5b6eg25FhGcYfxuD5P7TREuvVqdcJo5GzNj1LcslbnvmBsak7LZ4n7wNAbbCvXZPiOMnY1J2WzxP3mZaLM+hrEndO16alUyRuQRuCOYy/GU3cd9hIyOYy/GU3cd9ghZR8AFQPuGKSZ6Mijc9y7zUzUy7RbpLhNki6iJvu3/YnKS+jpIKSLW4I0am+u+vdUCMQWCvkRFfrcXI523+o91w3UZbVTFn3FJMCZrk86ZsraeNsyo6RGojlTcVeE+LjAlTQzQZZq5q5d3e/We4IrngMi4sSKvqI03GyORO5mY5ycU1w/wBZ6f4q/SpnGDh/rPT/ABV+lTOOKufz/j5PjL9J8H3P+Pk+Mv0nwckD1paeaqlSKCNXuXg3u6KWCSpqGQRJm965ITW3UUNDTpFEm3752+5QNTSYbYiI6qnVV/Nj2k8Kma2x2xEyWBXcqvd6TZHxLNDF+MljZ8ZyIRWtmsFvemTGSRrwteq/Tmaa6WSopGrLGuvRJuqibbe6hLIpYpUzikY9P0XIp9Ac8NvRWGrqGJJI5sDV3Edtu8Bt47NAy6rVIia0iapsfA70G1GZk0DcNR5e2q3qvIzL7TNtFqbb5nyNnWRHtyyVuWRnyzRRZa7Kxme5qnIh+xyRyfi5Gv8AirmB9EUxb1zb8kn0qSsimLeubfkk+lRBLTgAqAPuGKWaRI4o3Pcu81MzZwWCvkTN6RxfGdt/qA1JKMH/AMSm+U+xDGTDUuW3VMz+IptbLQOt8D43SJJqnarNEy3iKzyFX3rvU/G+wmporhYX1VZLUJUtaj1zyVu4CUZBv9jUnZbPE/eNjUnZbPE/eXNHtg78TU/Gb9Cm+NfZLc63sla6VJNWqLtJllkbAiobiTr1Uf3fNQ1xscSdeqj+75qGuKgAAAAAAAAAABM8PTOmtMSuXNzc2Z9zc/VkQwmmHoXQWmFHJk52b1Tu7n6siSsM6RjZI3RuTNrkVF7inPnJqXK1d5cjoL3Ixjnu3Gpmpz92qc5VVF21zEEvkH7kvAoyXgUqOgRfimfFQxb31pqfiKZUX4pnxUMW99aan4ikVCAAVHpT/j4/jJ9JPyAU/wCPj+Mn0k/JKw12Jes03db5yENJliXrNN3W+chDRBIACoAz6W0V9QiObArGrvvXL95nR4bqFT29RE1eRFUDAsPXen+N9ik1NHb7E+lrYqhalrkYueSNyz2jeElYaPGH8Ug+U+wjBNL1b3XCGONsqR6l2eapnvGq2NSdls8T94GgMuzddab5RDabGpOy2eJ+89qKwSU9XFOtS1yMcjstTulzRvjTYv62x/LJ5rjcmmxf1sj+WT6FIsoqACoHvR0dRWSainjVypurvJ3VPS1UT66rSFq5NTbe7gQmlLBFTQthhYjWJ+sg0VNhvaRamo7rY0+1fQZiWC3omWpkXl1ZtTymqaeFcpZ4o14HPRArA/AFu/Mk8czKCjhooligRyNV2qXNc9v/ANofP4RoOzIPHQfhGg7Mg8dAMo5/N+Nf8ZSbfhGg7Mg8dCESqiyOVNxVUQS+QAVAAAAAAJrYOs9P8VfpUhRNbB1np/ir9KklYZxzw6Gc8EEgB7UVPJV1LII09s5d3gThKhR0s9XLrcEavdv8Cd039JhyNERaqdzl/NZtJ4VNvQUkNHTpDC3JE3V33Lwqe5M1ya5tjtiJktOruVXu9J5zWC3vT2jZIl/Rdn9JsZJ4YlykmjYv6TkQ+o5I5UzjkY9OFq5gRG6WaoomrI1dehTdcibad1DWJtrkh0NURUyVM0NVS2aGC6OqkRFjRM42fmu3/wB37hmZNTR4fq5mI+Z7YEXeVM3eAzUw1Hlt1b1XkZ+83x8SzQxKiSyxsz3NU5EGZkwrRa0tz5VbMsiSIm0rcsss/SbA+Y5I5Ezje16formfRBEcU9dnfEaao2uKeuzviNNUckCT4P8A4nN8p9hGCT4P/ic3yn2CSG8IVfuu9R8b7CakKv3Xeo+N9hIWWCACoA9qamnqX6iCJ0i7+SbhsocPVz0zkdFHyK7Nf1AacmWGus0H97zlNZsal7LZ4i+k3drpVo6GOnV6PVme2iZbqqpJWGSc/n/HyfGX6ToBHZMNyOe53TTEzVV9x+8QSjwN/sak7LZ4n7xsak7LZ4n7y5o2GFutDPju+k2hiWmkWio0p3PR6oqrmiZbplkVA6/+PVHyrvpPA96/+PVHyrvpU8CoHrTQTVEqRQRue9d5D9o6eSqqWQRJm5y+DlJpb6KGigSKJu375y7rlA0lJhuRyI6pnRn6LEzXwmc3D9vRuSpK5eFXm2PiWeGH8bNHH8ZyIRWu/AFu/Mk8cyaC301Cr1gRyK/LPN2e4fv4RoOzIPHQfhGg7Mg8dAMohF767VPx1Jb+EaDsyDx0Ijd3skudQ+NyOar80VF2lEEsQAFQJTZrJFFG2arYj5V20Yu430qaWwQpNdoGuTNrVVy/Mmf0k0JKwIiImSJkiA117uaW+NiNYj5H55Iq7SIm+pon3+4OXNHRt5EZ6QJcCIfh65f0rPEQfh64/wBIzxEGRml5gYh6zVHcTzkNB+H7j+fH4iHlVXitqad8ErmKx27k3IZGbXgAqP1rVc5GtRVcq5Iib5MbNa4qGJHORHVCp7Z3ByIaDDUSS3aNXJmjEV/o/WpMCSsANRiK5yUSMggySV6Zq5Uz1KEbkrauR2qfUzKvx1GRmnYIxhWaaS4vbJK96a0q5Ocq76EnAAjmLJZY6qFI5XsRWbepcqb5pemansiXx1GRmnoIF0zU9kS+Ops8MTTSXRGvlkcmoXaVyqMjNKiOYy93S9x32EjI5jL3dL3HfYIJR8/WNV70Y1M3OXJEPwzLK1H3WmRf6RF8G2VEwt9KyjpGQM96ntl4V31PZzmtarnKiNRM1Vd4/TS4tndHRRwtXLXXe27ib30EVj3LEKo5Y6FqZJ/3jk+hPSamS5V73ZurJkX9F6p9BiAqJVhWeeenmWaV8mT0RFcueW0bk0+EWK23Pevv5Fy7mSfvNwRUIvXXap+OphmRcX65cKh6biyuVPCY5UTXD/Wen+Kv0qZxg4f6z0/xV+lTOOKufz/j5PjL9J8H3P8Aj5PjL9J8HJEhwhTIuvVbk209o36V+wkRrcMsRlniXfcrlXwqn2GyXaTMitBiO7SRSLSUrtS5E/hHpupyIRxyq5VVyqqruqp91EizVEkrt17lcvznmVH3DLJDIkkT3Mem4qLkTCxXDp+lVX5JNHtPy3+BSGG1wvMsd1azPalarV+n7CSJcYl4rOkaF8yZK9fasReFTLI9jJ65U0ee17Zy/q/eFaCaSSaRZJXq97t1VU3WDv41P8RPpNEb3B38an+In0llEmIpi3rm35JPpUlZFMW9c2/JJ9KkhZac96CmkrKpkEe65dteBN9TwJDg6JM6idU29pifSv2FRuqCjgooUihZlwu33LymQCN368TtqX0tK/W2s2nOTdVfsIqSAgS1VUq5rUzKvx1JJhOSSSjmWR7nqkm65c95BkZtyAQ691E7brUNbNI1EdtIjlRNwCYggXTNT2RL46jpmp7Il8dRkZp6DR4SkkkhqFkkc/JyZapc+E3gVDcSdep/7vmoa42OJOvU/wDd81DXFcQAAAABm2aGCor2wVCKrZEVEVFyVF3UU2NXhyVua007Xp+a/aXwmlp5Vhnjmbuscjk+ZSfMcj2Ne1c0cmaEVCpbVcI1ydSSL8VNV9AjtVxeuSUkifG2vpJsBmZI9a8Pq2Rsta5qom2kbdvPuqSEGvud2pqJqtRySzbzGrud3gA88S1jaegdEi/wkyalE5N9TDw/eNVqaSrdt7kb13+RTRVlTLVzummdm5fAicCHiMkdDBGbbf3QU6RVMbpVbtNci7eXKZWySn7Gl8KBc28MO99aan4imW1dU1HcKZmJe+tNT8RQIQACo9Kf8fH8ZPpJ+QCn/Hx/GT6SfklYa7EvWabut85CGkyxL1mm7rfOQhogkJTh61MhiZVVDEdM5M2ovvE9JHrbEk9fBE5M2uemfc3ydiSAGBfK9aClRzERZXrk3PcTlIpNX1sztVJVSqvAjsk8CAToEOsc877rTtdNI5FdtorlVNxSYgAaXFkkkdLCsb3MVX+9XLeI50zU9kS+OoM09BAumansiXx1Mq0VE7rnTtdPKqLImaK9RkZpmafF3WyP5ZPoU3Bp8XdbGfLJ9CglFAAVEtwrTpFbteVPbTOz+ZNpPt8JtjGtTUbbKZE/omr+o9ql6x08kibrWK7wIRUcxBd5XTPpaV6sYxcnuTdcu+ncNEu2uan6qqq5rtqfhUAAAAPrUP1tZNSuoRclXezA+QAAAAAAACa2DrPT/FX6VIUTWwdZ6f4q/SpJWGcc8OhnPBBISTCFMiRS1Tk23LqG9zdX/wB8hGyaYfYjLPTpworvCqiSGeR3EV2kbM6kpXqxG7Uj03c+BCQyORkbnruNRVOfyPdJI57lzc5VVV5RBL8VVVc1XNVPuCaWCRJIZHMem4qKeYKia2WvSvpNWqIkjV1L0Th4TOIphOZWXF0We1KxdrlTb9JKyKwb5WrQ0KyMy1xy6lnd4SGSPfI9XyOVzl21VVzVTe4xeqz08e8jVd4V/caAQkpBg38ZU9xv2kjI5g38ZU9xv2kjErCI4p67O+I01RtcU9dnfEaaoqBJ8H/xOb5T7CMEnwf/ABOb5T7BJDeEKv3Xeo+N9hNSFX7rvUfG+wkLLBM2zUDq+q1vNWxt23uTg4O6YRLMJxIy2rLltyPVc+RNr0lRs6eCKniSKFiMYm4iHoCJXS9VM8zmU8jooUXJup2lXlzIqWggXTVT2RN46ktw65z7RC57lc5dVmqrmvulA2ABBJqmoSZ6JUS+6X36gTsEC6ZqeyJfHUdM1PZEvjqMjNPQa3DT3vtTXPc5y6p22q5rumyIqB3D+P1HyrvpU8D3uH8fqPlXfSp4HJxSPB9OmomqlTbz1DfpX7CQGsww1Es8ap75zlXw5fYbMitFiK7Pgf0pTO1MmXt3put5E5SNOc5zlc5Vcq7qqu2p610iy1s0i7rpFX9Z4lQAAAA+mse5rnNaqoxM3Km8gHyAANjhyRI7vDqlyR2bfCm0TI561ytcjmqqORc0VN4mFmusVbG1j3IyoRNtq++5UJKw8sR22atbHLT5K+PNFaq5ZoRx9BXMXJ1JP8zFVCdAZmSCNoa125STr/u1PZlouT9ylendVE+kmoGZkh/4CuWWest7mrQxaqhq6VM56d7E/O3U8KE6CoioqKiKi7qKMzJzwG/xDaGxMdV0rcmJ7tib3KhoCo2+E3Il1VF99GqJ+pSWEEttR0rXRT7zXbfc3F/UTpjmvYj2qitcmaKm+hJWEZxhG5K2KXL2ro9Tnyoq+lDRk8rqSCsgWGduabqKm6i8KGndhqPVe1q3I3gVma/SMzJi4QRVuMjt5Il+lCUmLbaCCgiVkKKqu90526plARjGH8chT4P7TRm0xPO2a6uRq5pG1GfPur9JqyoG2wr12T4jjUm2wr12T4jgJaR3GXuqXuO+wkRHcZe6pe4/7CQso8ZNqlSG408jlyRJEzXgQxgVHQzWYion1lEixJnJEuqROFN9D5w/c21cDYJXZTsTLb98nCbUiueuRWqqKioqbqKelLTy1MzYYWK5y/q5VJxNSUs7tVNTxPdwuaiqfcMMMLdTDEyNOBrUQZmT4oKdtJRx07Vz1CZKvCu+p+V86U1FLOq+4aqp3d79Z7kYxPcWzPSkhdmxi5vVN9eD5gNGAComuH+s9P8AFX6VM4wcP9Z6f4q/SpnHFXP5/wAfJ8ZfpPg+5/x8nxl+k+DkiaYeXOzU/cXzlM2RFWNyJuqimqwnKj7asWe3G9Uy5F2/SbciueAy7xTLSXCWJUyaq6pncXcMQqB70E6U1ZFOrVcjHZ5JvngAJLslh7Ff4yGqvdwbcJY3tjViMaqZKuZrwAN7g7+NT/ET6TRG9wd/Gp/iJ9IkSYimLeubfkk+lSVkUxb1zb8kn0qSFlpyTYOci01Q3fR6L+r9xGTbYXqkp6/WnrkyZNT8+9/75SolpCL3G6O61DXJuvVydxdsm5hXO2U9eiLJmyRqZI9u73CKhJKcIIqUErt5ZfsQ82YajR+b6t6t4EZkvhzN1SwRU0DYYW6ljdxAPQhN927vU/H+wmrnI1qucqIiJmqrvEDrZdfrJpk3HvVydzMQS8QAVElwd+JqfjN+hTfGhwd+JqfjN+hTfElUNxL15n/u+ahrjZYl68z/AN3zUNaVAAAAAAMqC4VsKIkdVKiJtImqzRPmUxQB0CB2rgjfnnqmov6jAxHVVFJRMlpn6hyyI1VyRdrJeE9bFMk9qgdnttbqF+baP280y1duliamb8tU3uoRUTnuNdOmUlVIqLuoi5J+oxAu0uSgqAAAAADoMX4pnxUMW99aan4imVF+KZ8VDFvfWmp+IpFQgAFR6U/4+P4yfST8gFP+Pj+Mn0k/JKw12Jes03db5yENJliXrNN3W+chDRBLMsjkbdqZV/PRPDtE3OfRvdHI2Rq5OaqKndQnlHUMqqaOeNfavTPuLwCSGnxjG5aeCVE2muVF+dP3EZOgVEMdRC6GVqOY5MlRTSS4aiV6rHVPa3gc3P0AanD6K68U+W8qr+pSaGBa7VT0Cq9iukkVMle7eTkM8DR4x/ikCfpr9BGDfYwna6aGnau2xFc759z6DQlhAy7N11pvlEMQy7N11pvlEAnBp8XdbGfKp9Cm4NPi7rYz5VPoUiyigAKicWd6PtVM5N6NE8G19hkTs1yF8f5zVb4UNPhKpR9K+lcvto11TU5F/f8ASbsiuevarHK1yZORclTgPwkWIbRI+V1XSt1Su25GJu58KEeVFRclTJUKj8AAAl1kt7G2jWqhmev+2c1d5N419isz3yNqaxitYm21i7ru7yElJKwhN3t8lBUalc3RO9w/h5O6YRObq2ldQyJWKiRIm7vou9lykHXLNcs8t7MqPwAAAAAJrYOs9P8AFX6VIUTWwdZ6f4q/SpJWGcc8OhnPBBITeyLnaabL8xCEEvwvKklqazPbjcrV+n7RJDPrEVaSZE3Vjd9BAToa7aZEFuVM6krZIHJtNX2vKm8IJYwAKjJtlSlJXR1CtVyMz2kXdzRU+03myWHsV/jIRoAZ96r23CpZK2NWI1mpyVc99fSYAAEgwb+Mqe437SRkcwb+Mqe437SRklYRHFPXZ3xGmqNrinrs74jTVFQJPg/+JzfKfYRgk+D/AOJzfKfYJIbwhV+671HxvsJqQq/dd6j432EhZYJMMMORbPGie9c5F8OZDyQYRqmtdJRuXJXLq2cq7/2FlISNdtMjn8rHRSvjcmTmOVq91DoBrbnZqatkWXVOilXdc3cXuoRZQ4meHEVtmgz39Uv7SmFBhuBr0WaofI1Peo3U5m8YxrGNYxqNa1MkRN5BI/Tn823M9f0lJ3WTtpqWWdy5Ixqr8+8QEQSAAqJfhbrQz47vpNoavC3Whnx3fSbQioHceuFR8q76VPAyLj1wqflXfSpjlRL8LPR1pa1PePci/T9ptCNYRqUZNJSuXLVpqm91N39X0ElIqCXKJYa+eNUyyeuXc3jHJXiG1Oq8qinRNeamSt/OT0kWkY+N6se1WuTdRUyVCo+QAAJPhWj1NJLPK1FSb2qIqbrf3/Ya2zWiWre2WZqsp02812lf3PSS1jWsYjGoiNamSIm8hJWESv1rdRSa7EirTuXa/RXgU1R0CpSJ0D0n1OtaldVqtzIgdQkXTD0p1c6PVe0VU21QI8z9TNNtNo3dqsMkuUtZnGzdRie6Xu8BtbtbIp7brNPG1jovbRoib++nzlzMkeprxcIERqTq9qbz01X690ykxHW5bcUC/MvpNMqKi5LtKfgG4fiKuXcZA3uNX0ni6+XJ25Oje4xPQa0AbWC/3CNyLI5kqb6OaifQSa3VcdbStnjRUz2lRd5eAghKsItclvkcu46RcvAhJWG4e1r2q1yIrVTJUXfQgdZFrFXLD+Y9Wp4SekJviot2qcvzxBLCNvZLy6jRIJ0V8G8qbrf3GoBUT2mq6apbqoJmP5EXb8B7HPE2tw9UqJ0TJJpET4ykyXNPJJI426qR7WN4XLkhpbtfYmMdFRO1ci7Wr3m9zhUjLnOcubnKq8KqfgyM36qqqqqqqqu6qn4AVA2uFnIl2bmqJmxyJymqPqJ7o5GyMVWuauaLwKB0E0mLaZ8tNHUM20iVdUnIuW3+o2Vsq21tGyduSKu05OBd9DIe1r2Kx6IrXJkqLvoRXPQZl3onUNY6LbVi7bF4UMMqP1jnMcjmOVrkXNFRclQ3lBiKWNqMq49dRPft2neDfNEAJgy+21yZrK5nIrF+w+JcQUDE9ossi8jcvpIkCZLm21yvtTVNWOJNYjXdyXNy/OakAqAAAmeHHtfaIUauatza7kXM2BEcN13StbrT1yim2l5F3lJcRUGutM+krpIn8OqavCi75ikxxDQdOUmqjTOaLbbypvoQ4qNnh2tSkrtTIuUUvtXLwLvL/wC+EmBzwkNivTWsbTVjsstpki/QvpJKw2t3t0VwgRrl1MjfcP4OTuEVrLbW0rlSSBytT3zUzRfnJuioqIqKiou4qADnyMe5ckY5V4EQ2dtslVUvR0zXQRb6uTbXuIS4DMyQCphfT1EkMie2Y5UU8yQYupMnsrGJtO9o/u7ykfKgbzB7kSsmaqpmse0nDtmjPWknfTVLJ41ycxc+7yAT4jeLqWTXY6tNtmp1C8i7am/pJ2VNOyeNc2vTPuch+1MMdRA+GVM2PTJSKgAPevppKOqfBJutXaXhTeU8CokdovzdQ2GuVUVNpJd3Pu+k3sMsUzdVFIx7eFq5nPz9a5Wrm1VReFFJkuboR5VFRBTt1U0zI0/SUgy1E6pks0mXxlPNVVVzXbGRm3V8vXTLFpqXNIl909dpXcncNIAVAAASXB34mp+M36FN8aHB34mp+M36FN8SVQ7EvXmbuN81DWm6xNRVCV0lWkauhcie2TbyyRE2+A0pUAAAAAAAAbvCtakM7qSRcmSrm3kd+8lBz1FVFzRclJPZL0yZraerejZU2keu4795JWH1d7Gype6emckcq7bmr7ly/YpoZ7ZXwrk+lkXlamqT9RNwMzJAel6jPLWJc+DUKbO3WGpn9vU5wR8C+6X5t4lYGZk59Ix0cjo3Jk5qqipynybPEsGs3V6omTZUR6fb+s1hUT+lkZLTRyMXVNc1FRT5roOmaOWBHalXtVEXgNFhSuyctDIu0vto+7vp9vhJGRXP5o3wyuikarXtXJUPgk2KbfrkfTsTfbsTKRE304fmIyVH3CqJMxVXJEcir4ToCKipmi5oc8JVheu1+mWlkd/CRJ7Xlb+70ElYZ14p31Vumgj92qIqcqoueRCFRUVUVFRU3UU6ERnFNBrUvTkTfaPXJ6JvO4fnEEtEbCz3SWgerVRXwuX2zODlQ14KidUdwpKtqLDM1XL71Vyd4DJOeHo2eZqZNmkRORykyXNPnOa1NU5URE31U1dyvlLTsVsDknl3tT7lO6pE3ve9c3vc7urmfIyM33PLJNM6WVyue5c1VT4AKgZVpcjbnTOcqImuJtr3TFAHQzXYipn1Vtc2Pbcx2ry4cs/SMP13TlEiPXOWP2r+XgU2JFc8BtcR0HSlXrsbcoZVzTkXfQ1RUe9DUyUlSyeJdtq7abypwE0oKuGtp0mhd8Zu+1eBSCHtSVM9LKksEisdv8C90CemPVUNJUrnPTsev52WS+FDVUWIoXojauNY3fnN208G79JtIbhRTJ/B1US8iuyXwKRWKtitueesu7mrUyqW30VMucNOxrk3HLtr4VPdJY1TNJGeMeclXSxpnJUwt7r0A9j5mljhidLK5GsamaqprKu/UMKKkbnTu4GpknhUj90ulRXrqX5MiRc0Y3c+fhBm/bzcpK+ffbC1faN+1eUwACoAAAAABNMPOR1ngyVFyRUXwqQs3GGK7peq6WkX+DlXa5Hfv9BJISsgdfTSUdU+CTdau0vCm8pPDVYjoOm6XXo25zRJmn6Sb6BZRE22Ga1KatWKRco5skz4HbxqQVHQzAvFtjuESbaMmb7l/wBi8hr7HemqxtNWv1Lk2myLuL3fSb9FRUzTbQioNV2+spnKksD8k981M08JjtjkcuTWOVeBEOggZmSJW6x1VSuqnRYI/wBJPbL8xrJo3wyviemTmKqKh0AjOLaTUVDKtie1k9q/4yfu+gGTRAAqN/g1yJNUtzTNWtVE8PpJIQOhqX0lUyePdau2nCm+hOaeVk8LJo1zY9M0UkrCN4tppG1TardjeiN7ioaMn1ZTx1VM+CRPavTLucpBqunkpal8EqZOYuXd5RBLyJNg5ydKztzTNHouXzEZMy0VjqGtZLt6hfavThQqJuRHE1NJDcXTLtsm22ry76EtY5r2o5qorVTNFTfQx7nRsraN8DskXdY7gUioKfUb3xyNkjcrXNXNFTeP2WN8Uro5Gq17VyVOU+ColVrv0EzUjq1SKX873rvQbhj2PajmOa5q7iouaHPT6Y97FzY5zV5FyJkuboJ4VVbS0rVWedjF4M9vwEHdPO5MnTSKnK5TzGRm2l7urq9yRxorIGrmiLuuXhU1YBUAABLsKuRbUiIqKqPci8htSHYeruk61GvXKKX2ruRd5SYklULvtM+muUur22yOV7V4UVTAJpfKFK6iVGp/Cs9sxfs+chioqLkqZKhUfUUj4pWyxuVr2rmioTO0XGKvgzRUbK1Pbs4OVOQhJ9wyyQyJJE9WPbuKigdAPGppaapTKeFknAqptp85pKDESZIysiXP89n2obaC50Eye0qo+45dSv6yK8HWK2quaQuTkR6ntT2uggVHR0zNUm+7230mSksSpmkjFTkch8vqaeNM3zxN7r0QD1DlRqKqqiIm2qrvGtqb3b4UXKVZXcDEz/XuGiul5qK1qxNRIoV3Wouar3VBm+7/AHVat6wQKqU7V21/PXh7hqACol2HLh03S6zI7OaJMl/STeU2pAqKpkpKls8S+2bvLuKnAS+23WlrWojXJHLvxuXb+bhIrwutkhq3LLE7WZl3dr2ru6aKostxhVf4DXE4WLn+8mQBkgi0Vai5LSTp/u1PqO3V71ybSTfOxU+knIGZki9Dh6oe5HVbkiZvtRc3L9hJYIo4IWxRNRrGpkiH2fE0sUMaySyNY1N1XLkAnlZDC+WRcmsRVUgU8jpp3yu3XuVy/ObS+3Za1dYgzbAi7aruvX0GoEEgAKgAAAAAAAAAAM+0XOS3ufkzXGPTbaq5bfCbLZM7sNOc/cR4AbS7XZtwgax1KjHNXNr0fnlw7xqwAAAAAAAAAAAAG9p8RyxwMZJTpI5qZK7V5Z/qNEAJDsmd2GnOfuNJWzMqKl8zItaR65q1FzyXfPEAAABlUdwrKTagmcjfzV208BsosSVKJ/CU8Tu4qp6TRgCQLiV+W1SNTuv/AHHhNiKtcmUbIo+VEVV/WaYAe9VWVVUuc8738irteDcPAAAAANlaLvJb43x62krHLmiK7LJTP2TO7DTnP3EeAGxu9ybcEYq0yRvZuOR+eacG4a4AAAAAAAAAAAAN7haspqdJo55UjV6ordVueEkzVRzUc1UVF3FQ54ZFJW1VI7OCZzE4N1F+YmS5p2ai52KnqM5KfKCXgRPar828Y1FiNFybVw5fpx+g3VLV01U3VQTNfwoi7afMBCq2jqKOTUTxq3gXeXuKY50GaKOaNY5WNexd1FTMj9zw8qZyULs/g3L9C+kZmSPA+pY3xSLHIxzHJuoqZKfJUAAAAAGfRXatpERrJdWxNxr9tPSbOLEu1lLSbfC1/wBhHQBJ9klNl/F5s/mMapxJM5FSngZHyuXVKaEDIetVUz1UmuTyOkdy73cPIAD6ie6ORsjFVrmrmi8Cm9TEsuW3Ssz+MpoABvlxJIqKi0jFRd1Fd+40cqtdI5zG6hqrmjc88uQ+QAPaiqJKSqZPHutXc4U30PEASHZM7sNOc/cedRiBJ4XwyUKKx6ZKmufuNEAAAAAAAAAAAAAADKtlbJQ1STMTVJlk5ueWaG32TO7DTnP3EeAG5uF7bWUr4JKNERdxdc3F3l3DTAAAAAAAAAAAAAAAAAAAAACbS5oABvocSSsiY19Mj3ImSu1eWfLuH3smd2GnOfuI8BkPWrlZNUvlji1prlz1KLnkeQAAy6O5VlImphmXUfmu20MQAb2PElQifwlPE5f0VVPSfa4lfvUjU7r/ANxHwBuJsQ1z0yY2KPlRua/rNdVVdTUuznmfJwIq7SfMeAAAAAbS03iSggdCsSSszzbm7LU8JqwBIdkzuw05z9xrbvcGXBzH9LJFI3a1SOzzTg3DAAAAAbe23yWkpWwOhSVGr7VVdlknBuGVsmd2GnOfuI8AMy61jK6oSdIEiflk7J2eq4N4wwAAAAAAAAAAAAG7pMQyw0zIn06SqxMtVq8s/wBRpABIdkzuw05z9xprhUMqqp87Ida1e2rUXPb4THAAAAAAAAAAAAAAAAAGbTXWvp0RGVDlam872yfrM6PEdWnu4YXdzNPtNIAJAmJX79G3nP3H47EsvvaVid1yqaADIbabEFfImTFji+K3b/Xma6onmqH6uaV8i/pLnkeQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB9Mc5jkcxytcm4qLkqHyANtRX6sgybNlOz9Lad4TeUV6oanJqyay9feybX69whoGQkGMIvb09Qm+itVf1p9KkfPbXJFpFjV7lYj0VGqu0m0p4gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAf//Z" alt="Birla Uttam" style={{width:"100%",display:"block",clipPath:"inset(0 0 27% 0)"}}/></div><div style={{fontSize:10,color:"rgba(255,255,255,.6)",fontWeight:600,letterSpacing:.3,textAlign:"center",marginTop:2}}>Mangalam Cement Limited</div></div>
        <div className="hdr-sep" />
        <div className="hdr-search">🔍 Quick search…</div>
        {pendingCount > 0 && <div className="notif" onClick={() => { setPage("requests"); setActiveTab("pending"); }}>🔔<div className="notif-dot" /></div>}
        <div style={{ display: "flex", alignItems: "center", gap: 10, borderLeft: "1.5px solid var(--border)", paddingLeft: 16 }}>
          <div className="hdr-avatar">{user.name.charAt(0)}</div>
          <div className="hdr-user"><div className="hdr-uname">{user.name}</div><div className="hdr-urole">{user.role}</div></div>
          <button className="logout-btn" onClick={() => { setUser(null); setPage("dashboard"); }}>↗ Sign out</button>
        </div>
      </header>
      <div className="layout">
        <aside className="sidebar">
          {Object.entries(groups).map(([gk, gl]) => {
            const items = NAV.filter(n => n.group === gk);
            if (!items.length) return null;
            return (
              <div key={gk}>
                <div className="sb-section-label">{gl}</div>
                {items.map(item => (
                  <div key={item.key} className={`sb-item ${page === item.key ? "active" : ""}`} onClick={() => { setPage(item.key); if (item.key === "requests") setActiveTab("all"); }}>
                    <div className="sb-icon-wrap">{item.icon}</div>
                    <span className="sb-label">{item.label}</span>
                    {item.badge > 0 && <span className="sb-badge">{item.badge}</span>}
                  </div>
                ))}
              </div>
            );
          })}
          <div className="sb-bottom">
            <div style={{ padding: "8px 12px", borderRadius: 10, background: "var(--gold-f)", border: "1.5px solid var(--gold-200)" }}>
              <div style={{ fontSize: 10.5, fontWeight: 700, textTransform: "uppercase", letterSpacing: .6, color: "var(--gold-dk)", marginBottom: 4 }}>Logged in as</div>
              <div style={{ fontSize: 13, fontWeight: 700, color: "var(--ink)" }}>{user.name}</div>
              <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 1 }}>{user.role}</div>
            </div>
          </div>
        </aside>
        <main className="main">
          {page === "dashboard" && <Dashboard requests={requests} customers={customers} users={users} user={user} setPage={setPage} setActiveTab={setActiveTab} salesEntries={salesEntries} />}
          {page === "new-request" && <RequestForm user={user} users={users} customers={customers} freightMaster={freightMaster} priceMaster={priceMaster} tpcAgents={tpcAgents} modeMaster={modeMaster} unitSourceMaster={unitSourceMaster} sourceMaster={sourceMaster} storageLocationMaster={storageLocationMaster} opCommissionMaster={opCommissionMaster} onSubmit={handleNewRequest} onCancel={() => setPage("dashboard")} />}
          {page === "requests" && <RequestsPage requests={visibleRequests} user={user} users={users} freightMaster={freightMaster} salesEntries={salesEntries} setSalesEntries={setSalesEntries} onAction={handleAction} activeTab={activeTab} setActiveTab={setActiveTab} />}
          {page === "reports" && <ReportsPage requests={requests} salesEntries={salesEntries} />}
          {page === "sales-report" && <SalesReportPage requests={requests} salesEntries={salesEntries} user={user} />}
          {page === "sale-updation" && <SaleUpdationPage requests={requests} salesEntries={salesEntries} setSalesEntries={setSalesEntries} user={user} />}
          {page === "ncr-calculator" && <NCRCalculator user={user} />}
          {page === "customer-master" && user.role === "Admin" && <CustomerMasterPage customers={customers} setCustomers={setCustomers} users={users} />}
          {page === "admin-masters" && user.role === "Admin" && <AdminMasters tpcAgents={tpcAgents} setTpcAgents={setTpcAgents} users={users} setUsers={setUsers} customers={customers} priceMaster={priceMaster} setPriceMaster={setPriceMaster} freightMaster={freightMaster} setFreightMaster={setFreightMaster} locationMaster={locationMaster} setLocationMaster={setLocationMaster} copMaster={copMaster} setCopMaster={setCopMaster} packingMaster={packingMaster} setPackingMaster={setPackingMaster} plantMaster={plantMaster} setPlantMaster={setPlantMaster} modeMaster={modeMaster} setModeMaster={setModeMaster} unitSourceMaster={unitSourceMaster} setUnitSourceMaster={setUnitSourceMaster} sourceMaster={sourceMaster} setSourceMaster={setSourceMaster} storageLocationMaster={storageLocationMaster} setStorageLocationMaster={setStorageLocationMaster} opCommissionMaster={opCommissionMaster} setOpCommissionMaster={setOpCommissionMaster} />}
          {page === "audit" && <AuditLog requests={requests} users={users} />}
          {page.startsWith("adash-") && <AnalyticsDashboardRouter dashKey={page.replace("adash-","")} requests={requests} />}
        </main>
      </div>
    </>
  );
}
