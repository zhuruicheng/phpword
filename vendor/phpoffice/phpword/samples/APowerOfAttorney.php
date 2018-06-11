<?php
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\Style\Paragraph;
//生成用户委托书
include_once 'Sample_Header.php';

$languageEnGb = new \PhpOffice\PhpWord\Style\Language(\PhpOffice\PhpWord\Style\Language::EN_GB);
$phpWord = new \PhpOffice\PhpWord\PhpWord();
$phpWord->setDefaultParagraphStyle(
    array('spacing'=> 120)
);
$phpWord->getSettings()->setThemeFontLang($languageEnGb);

$fontStyleName = 'rStyle';
$phpWord->addFontStyle($fontStyleName, array('bold' => true, 'italic' => true, 'size' => 16, 'allCaps' => true, 'doubleStrikethrough' => true));

$paragraphStyleName = 'pStyle';
$phpWord->addParagraphStyle($paragraphStyleName, array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER, 'spaceAfter' => 100));
$phpWord->addTitleStyle(1, array('bold' => true), array('spaceAfter' => 240));
// New portrait section
$sectionStyle = array(
    'pageSizeW'=> \PhpOffice\PhpWord\Shared\Converter::cmToTwip(21),
    'pageSizeH'=> \PhpOffice\PhpWord\Shared\Converter::cmToTwip(29.7));
$section = $phpWord->addSection($sectionStyle);
$fontStyle['name'] = '宋体';
$fontStyle['size'] = 20;
$fontStyle['bold'] = true;
$fontStyle1['name'] = '宋体';
$fontStyle1['size'] = 10.5;
$fontStyle2['name'] = '宋体';
$fontStyle2['size'] = 12;
$fontStyle2_line['name'] = '宋体';
$fontStyle2_line['size'] = 12;
$fontStyle2_line['underline'] =  "solid";
$fontStyle3['name'] = '宋体';
$fontStyle3['size'] = 16;
if($_GET["type"] == 0){
    $section->addText('商标转让证明',$fontStyle,$paragraphStyleName);
    $section->addTextBreak(1);
    $section->addText('转让人：1    朱芮成',$fontStyle2);
    $section->addTextBreak(1);
    $section->addText('             2',$fontStyle2);
    $section->addTextBreak(1);
    $section->addText('             3',$fontStyle2);
    $section->addTextBreak(1);
    $section->addText('受让人：1    王然',$fontStyle2);
    $section->addTextBreak(1);
    $section->addText('             2',$fontStyle2);
    $section->addTextBreak(1);
    $section->addText('             3',$fontStyle2);
    $section->addTextBreak(1);
    $section->addText('      兹证明，上列双方经协商，一致同意将转让人的第xxxx号xxx商标、第xxxx号xxx商标等XX件商标转让给受让人。',$fontStyle2);
    $section->addTextBreak(3);
    $section->addText('转让人：',$fontStyle2);
    $section->addTextBreak(1);
    $section->addText('（章戳或签字）',$fontStyle2);
    $section->addTextBreak(1);
    $section->addText('       年       月       日',$fontStyle2);
    $section->addTextBreak(1);
    $section->addText('受让人：',$fontStyle2);
    $section->addTextBreak(1);
    $section->addText('（章戳或签字）',$fontStyle2);
    $section->addTextBreak(1);
    $section->addText('       年       月       日',$fontStyle2);
    $section->addTextBreak(1);
    $section->addText('备注：1、法人或其他组织应由法定代表人（或授权人）签字并加盖公章；个人应由本人亲笔签字。',$fontStyle2);
    $section->addTextBreak(1);
    $section->addText('2、涉及共有商标转让的，转让人或受让人排首位的为代表人，其他共有人依次排列。',$fontStyle2);
    $section->addTextBreak(1);
    $section->addText('3、涉及多件商标转让的，可以将全部注册号码一并列明在同一份文件中。',$fontStyle2);
    $section->addTextBreak(1);
}
if($_GET["type"] == 1){
    $section->addText('商 标 代 理 委 托 书',$fontStyle,$paragraphStyleName);
    $section->addTextBreak(1);
    $textrun = $section->addTextRun();
    $textrun->addText('      委托人',$fontStyle2);
    $textrun->addText(" ".$_GET["name"]." ", $fontStyle2_line);
    $textrun->addText("是", $fontStyle2);
    $textrun->addText(" 中 ", $fontStyle2_line);
    $textrun->addText("国国籍、依 ",$fontStyle2);
    $textrun->addText(" 中 ", $fontStyle2_line);
    $textrun->addText(" 国法律组成,现委托 ", $fontStyle2);
    $textrun->addText(" 苏州火龙知识产权代理有限公司 ", $fontStyle2_line);
    $textrun->addText(" 代理 ", $fontStyle2);
    $textrun->addText(" ".$_GET["tr_name"]." ", $fontStyle2_line);
    $textrun->addText("商标的如下“√”事宜。", $fontStyle2);
    $textrun = $section->addTextRun();
    $textrun->addText(" ",$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 0){
        $str = '☑商标注册申请';
    }else{
        $str = '□商标注册申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("                            ",$fontStyle1);
    if($_GET["num"] == 1){
        $str = '☑撤销成为商品/服务通用名称注册商标申请';
    }else{
        $str = '□撤销成为商品/服务通用名称注册商标申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 2){
        $str = '☑商标异议申请';
    }else{
        $str = '□商标异议申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("                            ",$fontStyle1);
    if($_GET["num"] == 3){
        $str = '☑撤销连续三年不使用注册商标提供证据';
    }else{
        $str = '□撤销连续三年不使用注册商标提供证据';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 4){
        $str = '☑商标异议答辩';
    }else{
        $str = '□商标异议答辩';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("                            ",$fontStyle1);
    if($_GET["num"] == 5){
        $str = '☑撤销成为商品/服务通用名称注册商标答辩';
    }else{
        $str = '□撤销成为商品/服务通用名称注册商标答辩';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 6){
        $str = '☑更正商标申请/注册事项申请';
    }else{
        $str = '□更正商标申请/注册事项申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("               ",$fontStyle1);
    if($_GET["num"] == 7){
        $str = '☑补发变更/转让/续展证明申请';
    }else{
        $str = '□补发变更/转让/续展证明申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 8){
        $str = '☑变更商标申请人/注册人名义/地址 变更集体';
    }else{
        $str = '□变更商标申请人/注册人名义/地址 变更集体';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText(" ",$fontStyle1);
    if($_GET["num"] == 9){
        $str = '☑补发商标注册证申请';
    }else{
        $str = '□补发商标注册证申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
        $str = '商标/证明商标管理规则/集体成员名单申请';
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("    ",$fontStyle1);
    if($_GET["num"] == 10){
        $str = '☑出具商标注册证明申请';
    }else{
        $str = '□出具商标注册证明申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 11){
        $str = '☑变更商标代理人/文件接收人申请';
    }else{
        $str = '□变更商标代理人/文件接收人申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("           ",$fontStyle1);
    if($_GET["num"] == 12){
        $str = '☑出具优先权证明文件申请';
    }else{
        $str = '□出具优先权证明文件申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 13){
        $str = '☑删减商品/服务项目申请';
    }else{
        $str = '□删减商品/服务项目申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("                   ",$fontStyle1);
    if($_GET["num"] == 14){
        $str = '☑撤回商标注册申请';
    }else{
        $str = '□撤回商标注册申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 15){
        $str = '☑商标续展注册申请';
    }else{
        $str = '□商标续展注册申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("                        ",$fontStyle1);
    if($_GET["num"] == 16){
        $str = '☑撤回商标异议申请';
    }else{
        $str = '□撤回商标异议申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 17){
        $str = '☑转让/移转申请/注册商标申请书';
    }else{
        $str = '□转让/移转申请/注册商标申请书';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("            ",$fontStyle1);
    if($_GET["num"] == 18){
        $str = '☑撤回变更商标申请人/注册人名义/地址 变更集';
    }else{
        $str = '□撤回变更商标申请人/注册人名义/地址 变更集';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 19){
        $str = '☑商标使用许可备案';
    }else{
        $str = '□商标使用许可备案';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("                        ",$fontStyle1);
    $str = '体商标/证明商标管理规则/集体成员名单申请';
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 20){
        $str = '☑变更许可人/被许可人名称备案';
    }else{
        $str = '□变更许可人/被许可人名称备案';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("             ",$fontStyle1);
    if($_GET["num"] == 21){
        $str = '☑撤回变更商标代理人/文件接收人申请';
    }else{
        $str = '□撤回变更商标代理人/文件接收人申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 22){
        $str = '☑商标使用许可提前终止备案';
    }else{
        $str = '□商标使用许可提前终止备案';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("                ",$fontStyle1);
    if($_GET["num"] == 23){
        $str = '☑撤回删减商品/服务项目申请';
    }else{
        $str = '□撤回删减商品/服务项目申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 24){
        $str = '☑商标专用权质权登记申请';
    }else{
        $str = '□商标专用权质权登记申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("                  ",$fontStyle1);
    if($_GET["num"] == 25){
        $str = '☑撤回商标续展注册申请';
    }else{
        $str = '□撤回商标续展注册申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 26){
        $str = '☑商标专用权质权登记事项变更申请';
    }else{
        $str = '□商标专用权质权登记事项变更申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("          ",$fontStyle1);
    if($_GET["num"] == 27){
        $str = '☑撤回转让/移转申请/注册商标申请';
    }else{
        $str = '□撤回转让/移转申请/注册商标申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 28){
        $str = '☑商标专用权质权登记期限延期申请';
    }else{
        $str = '□商标专用权质权登记期限延期申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("          ",$fontStyle1);
    if($_GET["num"] == 29){
        $str = '☑撤回商标使用许可备案';
    }else{
        $str = '□撤回商标使用许可备案';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 30){
        $str = '☑商标专用权质权登记证补发申请';
    }else{
        $str = '□商标专用权质权登记证补发申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("            ",$fontStyle1);
    if($_GET["num"] == 31){
        $str = '☑撤回商标注销申请';
    }else{
        $str = '□撤回商标注销申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 32){
        $str = '☑商标专用权质权登记注销申请';
    }else{
        $str = '□商标专用权质权登记注销申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("              ",$fontStyle1);
    if($_GET["num"] == 33){
        $str = '☑撤回撤销连续三年不使用注册商标申请';
    }else{
        $str = '□撤回撤销连续三年不使用注册商标申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 34){
        $str = '☑商标注销申请';
    }else{
        $str = '□商标注销申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("                            ",$fontStyle1);
    if($_GET["num"] == 35){
        $str = '☑撤回撤销成为商品/服务通用名称注册商标申请';
    }else{
        $str = '□撤回撤销成为商品/服务通用名称注册商标申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    if($_GET["num"] == 36){
        $str = '☑撤销连续三年不使用注册商标申请';
    }else{
        $str = '□撤销连续三年不使用注册商标申请';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun->addText("          ",$fontStyle1);
    if($_GET["num"] == 37){
        $str = '☑其他';
    }else{
        $str = '□其他';
    }
    $textrun->addText($str,$fontStyle1);
    $textrun = $section->addTextRun();
    $textrun->addText(" ",$fontStyle1);
    $textrun = $section->addTextRun();
    $textrun->addText("委托人地址 ", $fontStyle2);
    $c = "";
    if($_GET["c"] == "市辖区"||$_GET["c"] == "县"){
        $c = "";
    }else{
        $c = $_GET["c"];
    }
    $textrun->addText(" ".$_GET["p"].$c.$_GET["address"]." ", $fontStyle2_line);
    $textrun = $section->addTextRun();
    $textrun->addText("联  系  人 ", $fontStyle2);
    $textrun->addText(" ".$_GET["managers"]." ", $fontStyle2_line);
    $textrun->addText("                                委托人章戳（签字）", $fontStyle2);
    $textrun = $section->addTextRun();
    $textrun->addText("电      话 ", $fontStyle2);
    $textrun->addText(" ".$_GET["phone"]." ", $fontStyle2_line);
    $textrun = $section->addTextRun();
    $textrun->addText("邮 政 编 码", $fontStyle2);
    $textrun->addText(" ".$_GET["address_code"]." ", $fontStyle2_line);
    $textrun->addText("                                  ".date("Y年m月d")."日", $fontStyle2);
}
$filename = $_GET["id"];
foreach ($writers as $format => $extension) {
    if (null !== $extension) {
        $targetFile ="doc/".$filename.".".$extension;
        $phpWord->save($targetFile, $format);
    }
}
echo "success";
