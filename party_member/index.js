import fs from 'fs'
import path from 'path'
import XlsxTemplate from 'xlsx-template'
import XLSX from 'xlsx'
import _ from 'lodash'

import dataMapping from './data_mapping'

export default function run(filePath) {
    extractSourceFile(filePath)
}

/**
 * 提取党员信息列表
 * @param filePath
 */
function extractSourceFile(filePath) {
    let workbook = XLSX.readFile(filePath)
    if (!workbook) {
        console.log('readFile failed')
    }
    let firstSheetName = workbook.SheetNames[0]
    let memberListSheet = workbook.Sheets[firstSheetName]

    let memberInfoList = XLSX.utils.sheet_to_row_object_array(memberListSheet)

    let transferredMemberInfos = memberInfoList.map((memberInfo)=> {
        return translate(memberInfo)
    })

    let groups = _.groupBy(transferredMemberInfos, (m)=> {
        return m.homeAddress
    })

    for (let group in groups) {
        fillDestFile(group, groups[group])
    }
}

function translate(memberInfo) {
    let newMemberInfo = {}
    newMemberInfo.name = memberInfo["姓名"].trim()
    newMemberInfo.idcardNumber = memberInfo["身份证号码"].trim()
    newMemberInfo.gender = memberInfo["性别"].trim()
    newMemberInfo.nation = memberInfo["民族"].trim()
    newMemberInfo.native = memberInfo["籍贯"].trim()
    newMemberInfo.isTaiWan = memberInfo["是否台湾省籍"].trim()
    newMemberInfo.birthday = memberInfo["出生日期"].trim()
    //学历
    newMemberInfo.diploma = dataMapping.diploma[memberInfo["学历"]]

    newMemberInfo.partyType = memberInfo["人员类别"].trim()
    newMemberInfo.org = memberInfo["所在党组织"].trim()
    newMemberInfo.joinPartyTime = memberInfo["入党时间"].trim()
    newMemberInfo.formalTime = memberInfo["转正时间"].trim()
    //岗位
    newMemberInfo.career = dataMapping.career[memberInfo["工作岗位"]]

    newMemberInfo.workingTime = memberInfo["参加工作日期"].trim()
    newMemberInfo.homeAddress = memberInfo["家庭住址"].trim()

    let phoneNumber = memberInfo["联系电话"].trim()

    newMemberInfo.cellPhone = phoneNumber.length === 11 ? phoneNumber : ""
    newMemberInfo.distinctNumber = dataMapping.fixedValues.distinctNumber
    newMemberInfo.phone = dataMapping.fixedValues.phone
    newMemberInfo.marrigeStatus = memberInfo["婚姻状况"].trim()
    newMemberInfo.filePlace = memberInfo["党员档案所在单位"].trim()
    newMemberInfo.professionalTitle = memberInfo["技术职称"].trim()
    newMemberInfo.socialLevel = memberInfo["新社会阶层类型"].trim()
    newMemberInfo.isInFront = memberInfo["一线情况"].trim()
    newMemberInfo.trainging = memberInfo["培训情况"].trim()
    newMemberInfo.isFarmarWorker = memberInfo["是否农民工 "].trim()//奇葩,标题后面居然有个空格,只能这样了
    newMemberInfo.isLostPartyMember = memberInfo["是否失联党员"].trim()
    newMemberInfo.lostTime = ""
    newMemberInfo.partyStatus = newMemberInfo.isLostPartyMember === "否" ? "正常" : "停止党籍"
    newMemberInfo.infomationMatchRate = memberInfo["信息完整度(%)"].trim()
    newMemberInfo.isTravelPartyMember = "否"
    newMemberInfo.travelTo = ""

    return newMemberInfo
}

function fillDestFile(group, memberInfos) {
    let distPath = path.join(__dirname, 'dist', group)

    if (!fs.existsSync(distPath)) {
        fs.mkdirSync(distPath)
    }

    for (let memberInfo of memberInfos) {
        let filePath = path.join(distPath, memberInfo.name + ".xlsx")
        fs.readFile(path.join(__dirname, 'doc', 'template.xlsx'), function (err, data) {

            var template = new XlsxTemplate(data);

            var sheetNumber = 1;

            var values = memberInfo

            // Perform substitution
            template.substitute(sheetNumber, values);

            // Get binary data
            var outputData = template.generate();

            fs.writeFileSync(filePath, outputData, 'binary')
        });
    }
}