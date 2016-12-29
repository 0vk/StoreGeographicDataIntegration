if (typeof require !== 'undefined')
    XLSX = require('xlsx');

var provinceLib = [
    //直辖市
    ['北京市'],
    ['上海市'],
    ['天津市'],
    ['重庆市'],
    //华北地区
    ['河北省', '石家庄', '唐山', '秦皇岛', '邯郸', '邢台', '保定', '张家口', '承德', '沧州', '廊坊', '衡水'],
    ['山西省', '太原', '大同', '阳泉', '长治', '晋城', '朔州', '晋中', '运城', '忻州', '临汾', '吕梁'],
    ['内蒙古自治区', '呼和浩特', '包头', '乌海', '赤峰', '通辽', '鄂尔多斯', '呼伦贝尔', '巴彦淖尔', '乌兰察布', '兴安', '锡林郭勒', '阿拉善'],
    //东北地区
    ['辽宁省', '沈阳', '大连', '鞍山', '抚顺', '本溪', '丹东', '锦州', '营口', '阜新', '辽阳', '盘锦', '铁岭', '朝阳', '葫芦岛'],
    ['吉林省', '长春', '吉林', '四平', '辽源', '通化', '白山', '松原', '白城', '延边'],
    ['黑龙江', '哈尔滨', '齐齐哈尔', '鸡西', '鹤岗', '双鸭山', '大庆', '伊春', '佳木斯', '七台河', '牡丹江', '黑河', '绥化', '大兴安岭'],
    //华东地区
    ['江苏省', '南京', '无锡', '徐州', '常州', '苏州', '南通', '连云港', '淮安', '盐城', '扬州', '镇江', '泰州', '宿迁'],
    ['浙江省', '杭州', '宁波', '温州', '嘉兴', '湖州', '绍兴', '金华', '衢州', '舟山', '台州', '丽水'],
    ['安徽省', '合肥', '芜湖', '蚌埠', '淮南', '马鞍山', '淮北', '铜陵', '安庆', '黄山', '滁州', '阜阳', '宿州', '巢湖', '六安', '亳州', '池州', '宣城'],
    ['福建省', '福州', '厦门', '莆田', '三明', '泉州', '漳州', '南平', '龙岩', '宁德'],
    ['江西省', '南昌', '景德镇', '萍乡', '九江', '新余', '鹰潭', '赣州', '吉安', '宜春', '抚州', '上饶'],
    ['山东省', '济南', '青岛', '淄博', '枣庄', '东营', '烟台', '潍坊', '威海', '济宁', '泰安', '日照', '莱芜', '临沂', '德州', '聊城', '滨州', '菏泽'],
    //中南地区
    ['河南省', '郑州', '开封', '洛阳', '平顶山', '焦作', '鹤壁', '新乡', '安阳', '濮阳', '许昌', '漯河', '三门峡', '南阳', '商丘', '信阳', '周口', '驻马店'],
    ['湖北省', '武汉', '黄石', '襄樊', '十堰', '荆州', '宜昌', '荆门', '鄂州', '孝感', '咸宁', '随州', '恩施'],
    ['湖南省', '长沙', '株洲', '湘潭', '衡阳', '邵阳', '岳阳', '常德', '张家界', '益阳', '郴州', '永州', '怀化', '娄底', '湘西'],
    ['广东省', '广州', '深圳', '珠海', '汕头', '韶关', '佛山', '江门', '湛江', '茂名', '肇庆', '惠州', '梅州', '汕尾', '河源', '阳江', '清远', '东莞', '中山', '潮州', '揭阳', '云浮'],
    ['广西自治区', '南宁', '柳州', '桂林', '梧州', '北海', '防城港', '钦州', '贵港', '玉林', '百色', '贺州', '河池', '来宾', '崇左'],
    ['海南省', '海口', '三亚'],
    //西南地区
    ['四川省', '成都', '自贡', '攀枝花', '泸州', '德阳', '绵阳', '广元', '遂宁', '内江', '乐山', '南充', '宜宾', '广安', '达州', '眉山', '雅安', '巴中', '资阳', '阿坝', '甘孜', '凉山'],
    ['贵州省', '贵阳', '六盘水', '遵义', '安顺', '铜仁', '毕节', '黔西南', '黔东南', '黔南'],
    ['云南省', '昆明', '曲靖', '玉溪', '保山', '昭通', '丽江', '普洱', '临沧', '文山', '红河', '西双版纳', '楚雄', '大理', '德宏', '怒江', '迪庆'],
    ['西藏自治区', '拉萨', '昌都', '山南', '日喀则', '那曲', '阿里', '林芝'],
    //西北地区
    ['陕西省', '西安', '铜川', '宝鸡', '咸阳', '渭南', '延安', '汉中', '榆林', '安康', '商洛'],
    ['甘肃省', '兰州', '嘉峪关', '金昌', '白银', '天水', '武威', '张掖', '平凉', '酒泉', '庆阳', '定西', '陇南', '临夏', '甘南'],
    ['青海省', '西宁', '海东', '海北', '黄南', '海南', '果洛', '玉树', '海西'],
    ['宁夏自治区', '银川', '石嘴山', '吴忠', '固原', '中卫'],
    ['新疆自治区', '乌鲁木齐', '克拉玛依', '吐鲁番', '哈密', '和田', '阿克苏', '喀什', '克孜勒苏柯尔克孜', '巴音郭楞蒙古', '昌吉', '博尔塔拉蒙古', '伊犁哈萨克', '塔城', '阿勒泰'],
    //港澳台
    ['香港特别行政区'],
    ['澳门特别行政区'],
    ['台湾省', '台北', '高雄', '基隆', '台中', '台南', '新竹', '嘉义']
];

var workbook1, workbook2, workbook3, wb1_sheet1, wb2_sheet1, wb3_sheet1;

/*
 * 输出文件列名如下:
 *  '商户名称（poi_name)'
 *  '省份（province）'
 *  '城市（city）'
 *  '区县（town）'
 *  '详细地址(poi_address)'
 *  '电话(phone)'
 *  '经度/纬度'
 *  '一级行业'
 *  '二级行业'
 *  '渠道来源'
 *  '品牌'
 */
var temp_PoiName = '商户名称（poi_name)';
var temp_Province = '省份（province）';
var temp_City = '城市（city）';
var temp_Town = '区县（town）';
var temp_PoiAddress = '详细地址(poi_address)';
var temp_Phone = '电话(phone)';
var temp_Coordinates = '经度/纬度';
var temp_PrimaryIndustry = '一级行业';
var temp_SecondaryIndustry = '二级行业';
var temp_Source = '渠道来源';
var temp_Brand = '品牌';
var temp_Judge = '诊断';

var template =
    [
        temp_PoiName,
        temp_Province,
        temp_City,
        temp_Town,
        temp_PoiAddress,
        temp_Phone,
        temp_Coordinates,
        temp_PrimaryIndustry,
        temp_SecondaryIndustry,
        temp_Source,
        temp_Brand,
        temp_Judge
    ];
var output_template = [];

/*
 *  以下是「SBUX(CN)_MDM_Store List_20160711.xlsx」文件内sheet1中的列名
 *  'Legal Entity'
 *  'Location Number'
 *  'Responsibility Center'
 *  'Location Name'
 *  'Location Purpose'
 *  'Ownership Type'
 *  'Current Lifecycle Status'
 *  'Country Code (physical)'
 *  'Open Date'
 *  'Close Date'
 *  'Global Region'
 *  'Division'
 *  'Region'
 *  'Area'
 *  'District Number'
 *  'Address Line 1 (physical)'
 *  'Address Line 2 (physical)'
 *  'Address Line 3 (physical)'
 *  'City (physical)'
 *  'Country Subdivision/State (physical)'
 *  'County (physical)'
 *  'Proposed Open Date'
 */
var wb1_LegalEntity = 'Legal Entity';
var wb1_LocationNumber = 'Location Number';
var wb1_ResponsibilityCenter = 'Responsibility Center';
var wb1_LocationName = 'Location Name';
var wb1_LocationPurpose = 'Location Purpose';
var wb1_OwnershipType = 'Ownership Type';
var wb1_CurrentLifecycleStatus = 'Current Lifecycle Status';
var wb1_CountryCode_physical = 'Country Code (physical)';
var wb1_OpenDate = 'Open Date';
var wb1_CloseDate = 'Close Date';
var wb1_GlobalRegion = 'Global Region';
var wb1_Division = 'Division';
var wb1_Region = 'Region';
var wb1_Area = 'Area';
var wb1_DistrictNumber = 'District Number';
var wb1_AddressLine1_physical = 'Address Line 1 (physical)';
var wb1_AddressLine2_physical = 'Address Line 2 (physical)';
var wb1_AddressLine3_physical = 'Address Line 3 (physical)';
var wb1_City_physical = 'City (physical)';
var wb1_CountrySubdivisionOrState_physical = 'Country Subdivision/State (physical)';
var wb1_County_physical = 'County (physical)';
var wb1_ProposedOpenDate = 'Proposed Open Date';

/*
 *  以下是「CHINA OPEN JV STORES_27_Jun_2016.xlsx」文件内sheet1中的列名
 *  'Location Number'
 *  'Location Name'
 *  'Location Purpose'
 *  'Ownership Type'
 *  'Current Lifecycle Status'
 *  'Country Code (physical)'
 *  'Proposed Open Date'
 *  'Open Date'
 *  'Close Date'
 *  'Phone Number'
 *  'Address Line 1 (physical)'
 *  'Address Line 2 (physical)'
 *  'Address Line 3 (physical)'
 *  'City (physical)'
 *  'Country Subdivision/State (physical)'
 *  'Province	'
 *  'County (physical)'
 *  'Postal Code (physical)'
 *  'Latitude (physical)'
 *  'Longitude (physical)'
 */

var wb2_LocationNumber = 'Location Number';
var wb2_LocationName = 'Location Name';
var wb2_LocationPurpose = 'Location Purpose';
var wb2_OwnershipType = 'Ownership Type';
var wb2_CurrentLifecycleStatus = 'Current Lifecycle Status';
var wb2_CountryCode_physical = 'Country Code (physical)';
var wb2_ProposedOpenDate = 'Proposed Open Date';
var wb2_OpenDate = 'Open Date';
var wb2_CloseDate = 'Close Date';
var wb2_PhoneNumber = 'Phone Number';
var wb2_AddressLine1_physical = 'Address Line 1 (physical)';
var wb2_AddressLine2_physical = 'Address Line 2 (physical)';
var wb2_AddressLine3_physical = 'Address Line 3 (physical)';
var wb2_City_physical = 'City (physical)';
var wb2_CountrySubdivisionOrState_physical = 'Country Subdivision/State (physical)';
var wb2_Province = 'Province';
var wb2_County_physical = 'County (physical)';
var wb2_PostalCode_physical = 'Postal Code (physical)';
var wb2_Latitude_physical = 'Latitude (physical)';
var wb2_Longitude_physical = 'Longitude (physical)';

/*
 * 以下是「OPENCOCHINA_27_Jun_2016.xlsx」文件内sheet1中的列名
 *  'Location Number'
 *  'Location Name'
 *  'Location Purpose'
 *  'Ownership Type'
 *  'Current Lifecycle Status'
 *  'Country Code (physical)'
 *  'Proposed Open Date'
 *  'Open Date'
 *  'Close Date'
 *  'Legal Entity'
 *  'Phone Number'
 *  'Address Line 1 (physical)'
 *  'Address Line 2 (physical)'
 *  'Address Line 3 (physical)'
 *  'City (physical)'
 *  'Country Subdivision/State (physical)'
 *  'County (physical)'
 *  'Postal Code (physical)'
 *  'Latitude (physical)'
 *  'Longitude (physical)'
 */
var wb3_LocationNumber = 'Location Number';
var wb3_LocationName = 'Location Name';
var wb3_LocationPurpose = 'Location Purpose';
var wb3_OwnershipType = 'Ownership Type';
var wb3_CurrentLifecycleStatus = 'Current Lifecycle Status';
var wb3_CountryCode_physical = 'Country Code (physical)';
var wb3_ProposedOpenDate = 'Proposed Open Date';
var wb3_OpenDate = 'Open Date';
var wb3_CloseDate = 'Close Date';
var wb3_LegalEntity = 'Legal Entity';
var wb3_PhoneNumber = 'Phone Number';
var wb3_AddressLine1_physical = 'Address Line 1 (physical)';
var wb3_AddressLine2_physical = 'Address Line 2 (physical)';
var wb3_City_physical = 'City (physical)';
var wb3_CountrySubdivisionOrState_physical = 'Country Subdivision/State (physical)';
var wb3_County_physical = 'County (physical)';
var wb3_PostalCode_physical = 'Postal Code (physical)';
var wb3_Latitude_physical = 'Latitude (physical)';
var wb3_Longitude_physical = 'Longitude (physical)';


function excel2data() {
    try {
        workbook1 = XLSX.readFile('resources/wb1.xlsx');
        workbook2 = XLSX.readFile('resources/wb2.xlsx');
        workbook3 = XLSX.readFile('resources/wb3.xlsx');

        wb1_sheet1 = XLSX.utils.sheet_to_json(workbook1.Sheets[workbook1.SheetNames[0]]);
        wb2_sheet1 = XLSX.utils.sheet_to_json(workbook2.Sheets[workbook2.SheetNames[0]]);
        wb3_sheet1 = XLSX.utils.sheet_to_json(workbook3.Sheets[workbook3.SheetNames[0]]);

        // clear
        output_template = [];
        output_template.push(template);

        // wb1_sheet1.forEach(function (column) {
        //     output_template.push(makeData(0, column));
        //     // console.log(`
        //     // 'Legal Entity' : {%s}
        //     // 'Location Number' : {%s}
        //     // 'Responsibility Center' : {%s}
        //     // 'Location Name' : {%s}
        //     // 'Location Purpose' : {%s}
        //     // 'Ownership Type' : {%s}
        //     // 'Current Lifecycle Status' : {%s}
        //     // 'Country Code (physical)' : {%s}
        //     // 'Open Date' : {%s}
        //     // 'Close Date' : {%s}
        //     // 'Global Region' : {%s}
        //     // 'Division' : {%s}
        //     // 'Region' : {%s}
        //     // 'Area' : {%s}
        //     // 'District Number' : {%s}
        //     // 'Address Line 1 (physical) : {%s}
        //     // 'Address Line 2 (physical) : {%s}
        //     // 'Address Line 3 (physical) : {%s}
        //     // 'City (physical)' : {%s}
        //     // 'Country Subdivision/State : {%s}
        //     // 'County (physical)' : {%s}
        //     // 'Proposed Open Date' : {%s}
        //     // `,
        //     // column[wb1_LegalEntity],
        //     // column[wb1_LocationNumber],
        //     // column[wb1_ResponsibilityCenter],
        //     // column[wb1_LocationName],
        //     // column[wb1_LocationPurpose],
        //     // column[wb1_OwnershipType],
        //     // column[wb1_CurrentLifecycleStatus],
        //     // column[wb1_CountryCode_physical],
        //     // column[wb1_OpenDate],
        //     // column[wb1_CloseDate],
        //     // column[wb1_GlobalRegion],
        //     // column[wb1_Division],
        //     // column[wb1_Region],
        //     // column[wb1_Area],
        //     // column[wb1_DistrictNumber],
        //     // column[wb1_AddressLine1_physical],
        //     // column[wb1_AddressLine2_physical],
        //     // column[wb1_AddressLine3_physical],
        //     // column[wb1_City_physical],
        //     // column[wb1_CountrySubdivisionOrState_physical],
        //     // column[wb1_County_physical],
        //     // column[wb1_ProposedOpenDate]);
        // });

        wb2_sheet1.forEach(function (column) {
            output_template.push(makeData(1, column));
            // console.log(`
            // 'Location Number' : {%s}
            // 'Location Name' : {%s}
            // 'Location Purpose' : {%s}
            // 'Ownership Type' : {%s}
            // 'Current Lifecycle Status' : {%s}
            // 'Country Code (physical)' : {%s}
            // 'Proposed Open Date' : {%s}
            // 'Open Date' : {%s}
            // 'Close Date' : {%s}
            // 'Phone Number' : {%s}
            // 'Address Line 1 (physical)' : {%s}
            // 'Address Line 2 (physical)' : {%s}
            // 'Address Line 3 (physical)' : {%s}
            // 'City (physical)' : {%s}
            // 'Country Subdivision/State (physical)' : {%s}
            // 'Province	' : {%s}
            // 'County (physical)' : {%s}
            // 'Postal Code (physical)' : {%s}
            // 'Latitude (physical)' : {%s}
            // 'Longitude (physical)' : {%s}
            // `,
            //     column[wb2_LocationNumber],
            //     column[wb2_LocationName],
            //     column[wb2_LocationPurpose],
            //     column[wb2_OwnershipType],
            //     column[wb2_CurrentLifecycleStatus],
            //     column[wb2_CountryCode_physical],
            //     column[wb2_ProposedOpenDate],
            //     column[wb2_OpenDate],
            //     column[wb2_CloseDate],
            //     column[wb2_PhoneNumber],
            //     column[wb2_AddressLine1_physical],
            //     column[wb2_AddressLine2_physical],
            //     column[wb2_AddressLine3_physical],
            //     column[wb2_City_physical],
            //     column[wb2_CountrySubdivisionOrState_physical],
            //     column[wb2_Province],
            //     column[wb2_County_physical],
            //     column[wb2_PostalCode_physical],
            //     column[wb2_Latitude_physical],
            //     column[wb2_Longitude_physical]);
        });

        wb3_sheet1.forEach(function (column) {
            output_template.push(makeData(2, column));
            // console.log(`
            // 'Location Number' : {%s}
            // 'Location Name' : {%s}
            // 'Location Purpose' : {%s}
            // 'Ownership Type' : {%s}
            // 'Current Lifecycle Status' : {%s}
            // 'Country Code (physical)' : {%s}
            // 'Proposed Open Date' : {%s}
            // 'Open Date' : {%s}
            // 'Close Date' : {%s}
            // 'Legal Entity' : {%s}
            // 'Phone Number' : {%s}
            // 'Address Line 1 (physical)' : {%s}
            // 'Address Line 2 (physical)' : {%s}
            // 'Address Line 3 (physical)' : {%s}
            // 'City (physical)' : {%s}
            // 'Country Subdivision/State (physical)' : {%s}
            // 'County (physical)' : {%s}
            // 'Postal Code (physical)' : {%s}
            // 'Latitude (physical)' : {%s}
            // 'Longitude (physical)' : {%s}
            // `,
            //     column[wb3_LocationNumber],
            //     column[wb3_LocationName],
            //     column[wb3_LocationPurpose],
            //     column[wb3_OwnershipType],
            //     column[wb3_CurrentLifecycleStatus],
            //     column[wb3_CountryCode_physical],
            //     column[wb3_ProposedOpenDate],
            //     column[wb3_OpenDate],
            //     column[wb3_CloseDate],
            //     column[wb3_LegalEntity],
            //     column[wb3_PhoneNumber],
            //     column[wb3_AddressLine1_physical],
            //     column[wb3_AddressLine2_physical],
            //     column[wb3_City_physical],
            //     column[wb3_CountrySubdivisionOrState_physical],
            //     column[wb3_County_physical],
            //     column[wb3_PostalCode_physical],
            //     column[wb3_Latitude_physical],
            //     column[wb3_Longitude_physical]);
        });

        return 1;
    } catch (e) {
        //TBD
        console.log(e);
        return null;
    }
}

function makeData(type, col) {
    // wb1_sheet1.forEach(function (_col) {
    //     if (col[wb1_LocationNumber] === _col[wb2_LocationNumber]) {
    //         console.log(col[wb3_County_physical]);
    //     }
    // })
    var data;
    switch (type) {
        // case 1 : {
        //     data.push(col[wb1_LocationName]);
        //     data.push(col[wb1_County_physical]);
        //     data.push(col[wb1_City_physical]);
        //     data.push(col[wb1_AddressLine1_physical]);
        //     data.push(col[wb1_AddressLine2_physical]);
        //     data.push(null);
        //     data.push(null);
        // }
        //     break;
        case 1 : {
            // data.push(col[wb2_LocationName]);
            // data.push(col[wb2_Province]);
            // data.push(col[wb2_City_physical]);
            // data.push(col[wb2_AddressLine1_physical]);
            // data.push(col[wb2_AddressLine2_physical]);
            // data.push(col[wb2_PhoneNumber]);
            // data.push(col[wb2_Latitude_physical] + '/' + col[wb2_Longitude_physical]);
            data = op_wb2(col);
        }
            break;
        case 2 : {
            // data.push(col[wb3_LocationName]);
            // data.push(col[wb3_County_physical]);
            // data.push(col[wb3_City_physical]);
            // data.push(col[wb3_AddressLine1_physical]);
            // data.push(col[wb3_AddressLine2_physical]);
            // data.push(col[wb3_PhoneNumber]);
            // data.push(col[wb3_Latitude_physical] + '/' + col[wb3_Longitude_physical]);
            data = op_wb3(col);
        }
            break;
        default : {
            console.log('ERROR!');
            return data;
        }
    }
    return data;
}

function matchProvince(city) {
    var result;
    var key = city.replace('市', '');
    provinceLib.forEach(function (p) {
        p.forEach(function (c) {
            if (0 === c.search(key)) {
                result = p[0];
            }
            ;
        })
    });
    return result;
}

function op_wb2(col) {
    var result = new Array(12), judgeProvince = false;
    result[0] = col[wb2_LocationName];
    if (col[wb2_Province]) {
        judgeProvince = true;
        result[1] = col[wb2_Province];
    } else {
        if (col[wb2_City_physical]) {
            var _p = matchProvince(col[wb2_City_physical]);
            if (_p) {
                judgeProvince = true;
                result[1] = _p;
            }
        }
    }
    result[2] = col[wb2_City_physical];
    result[3] = col[wb2_AddressLine1_physical];
    result[4] = col[wb2_AddressLine2_physical];
    result[5] = col[wb2_PhoneNumber];
    result[6] = col[wb2_Latitude_physical] + '/' + col[wb2_Longitude_physical];
    result[7] = '休闲娱乐';
    result[8] = '咖啡酒吧';
    result[9] = undefined;
    result[10] = undefined;
    if (col[wb2_LocationName] &&
        judgeProvince &&
        col[wb2_City_physical] &&
        col[wb2_AddressLine1_physical] &&
        col[wb2_AddressLine2_physical] &&
        // col[wb2_PhoneNumber] &&
        col[wb2_Latitude_physical] &&
        col[wb2_Longitude_physical] &&
        ('0' !== col[wb2_Latitude_physical]) &&
        ('0' !== col[wb2_Longitude_physical])
    ) {
        result[11] = undefined;
    } else {
        result[11] = 'NG!';
    }
    return result;
}

function op_wb3(col) {
    var result = new Array(12);
    var judgeProvince = false;
    result[0] = col[wb3_LocationName];
    if (col[wb3_County_physical]) {
        judgeProvince = true;
        result[1] = col[wb3_County_physical];
    } else {
        if (col[wb3_City_physical]) {
            var _p = matchProvince(col[wb3_City_physical]);
            if (_p) {
                judgeProvince = true;
                result[1] = _p;
            }
        }
    }
    result[2] = col[wb3_City_physical];
    result[3] = col[wb3_AddressLine1_physical];
    result[4] = col[wb3_AddressLine2_physical];
    result[5] = col[wb3_PhoneNumber];
    result[6] = col[wb3_Latitude_physical] + '/' + col[wb3_Longitude_physical];
    result[7] = '休闲娱乐';
    result[8] = '咖啡酒吧';
    result[9] = null;
    result[10] = null;
    if (col[wb3_LocationName] &&
        judgeProvince &&
        col[wb3_City_physical] &&
        col[wb3_AddressLine1_physical] &&
        col[wb3_AddressLine2_physical] &&
        // col[wb3_PhoneNumber] &&
        col[wb3_Latitude_physical] &&
        col[wb3_Longitude_physical] &&
        ('0' !== col[wb3_Latitude_physical]) &&
        ('0' !== col[wb3_Longitude_physical])
    ) {
        result[11] = '';
    } else {
        result[11] = 'NG!';
    }
    return result;
}

function export2file() {
    /* output format determined by filename */

    var ws_name = 'Result';

    var wscols = [
        {wch: 25},
        {wch: 10},
        {wch: 10},
        {wch: 20},
        {wch: 50},
        {wch: 20},
        {wch: 25},
        {wch: 10},
        {wch: 10},
        {wch: 10},
        {wch: 10},
        {wch: 10}
    ];

    /* dummy workbook constructor */
    function Workbook() {
        if (!(this instanceof Workbook)) return new Workbook();
        this.SheetNames = [];
        this.Sheets = {};
    }

    var wb = new Workbook();


    /* TODO: date1904 logic */
    function datenum(v, date1904) {
        if (date1904) v += 1462;
        var epoch = Date.parse(v);
        return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
    }

    /* convert an array of arrays in JS to a CSF spreadsheet */
    function sheet_from_array_of_arrays(data, opts) {
        var ws = {};
        var range = {s: {c: 10000000, r: 10000000}, e: {c: 0, r: 0}};
        for (var R = 0; R != data.length; ++R) {
            for (var C = 0; C != data[R].length; ++C) {
                if (range.s.r > R) range.s.r = R;
                if (range.s.c > C) range.s.c = C;
                if (range.e.r < R) range.e.r = R;
                if (range.e.c < C) range.e.c = C;
                var cell = {v: data[R][C]};
                if (cell.v == null) continue;
                var cell_ref = XLSX.utils.encode_cell({c: C, r: R});

                /* TEST: proper cell types and value handling */
                if (typeof cell.v === 'number') cell.t = 'n';
                else if (typeof cell.v === 'boolean') cell.t = 'b';
                else if (cell.v instanceof Date) {
                    cell.t = 'n';
                    cell.z = XLSX.SSF._table[14];
                    cell.v = datenum(cell.v);
                }
                else cell.t = 's';
                ws[cell_ref] = cell;
            }
        }

        /* TEST: proper range */
        if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
        return ws;
    }

    var ws = sheet_from_array_of_arrays(output_template);

    /* TEST: add worksheet to workbook */
    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = ws;

    /* TEST: column widths */
    ws['!cols'] = wscols;

    /* write file */
    XLSX.writeFile(wb, 'resources/output.xlsx');
}

exports.excel2dta = excel2data;

exports.export2file = export2file;