import  {separate,summary} from './party_member'

if (process.env.type === "separate") {
    separate(__dirname + '/party_member/private_doc/member_list.xls')
} else if (process.env.type === "summary") {
    summary(__dirname + '/party_member/private_doc/member_list.xls')
}