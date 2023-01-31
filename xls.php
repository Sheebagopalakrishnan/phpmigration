<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

define('IN_SCRIPT',1);
define('CIB_PATH','./');

$cib_settings['server_local']=0;

if ($cib_settings['server_local']) { 
require(CIB_PATH . 'cib_settings.inc.php');
require(CIB_PATH . 'inc/database.inc.php');
require(CIB_PATH . 'inc/common.inc.php');
require_once(CIB_PATH . 'inc/clean_file_name.inc.php');
$save_path='C:/xampp/htdocs/cib';
}
else { 
// Get all the required files and functions
require('C:/inetpub/wwwroot/php-dev/cib/cib_settings.inc.php');
require('C:/inetpub/wwwroot/php-dev/cib/inc/database.inc.php');
require('C:/inetpub/wwwroot/php-dev/cib/inc/clean_file_name.inc.php');


//require('/var/www/html/cib/cib_settings.inc.php');
//require('/var/www/html/cib/inc/database.inc.php');
//require('/var/www/html/cib/inc/clean_file_name.inc.php');

//$save_path='/var/www/html/cib';
$save_path='C:/inetpub/wwwroot/php-dev/cib';


function cib_error($error,$showback=1) {
global $cib_settings, $hesklang;
?>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="3"><img src="<?php echo CIB_PATH; ?>img/headerleftsm.jpg" width="3" height="25" alt="" /></td>
<td class="headersm"><?php echo "error"; ?></td>
<td width="3"><img src="<?php echo CIB_PATH; ?>img/headerrightsm.jpg" width="3" height="25" alt="" /></td>
</tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="3">
<tr>
<td><span><a href="index.php">Main Page</a>
&gt; <?php echo "error"; ?></span></td>
</tr>
</table>

</td>
</tr>
<tr>
<td>
<p>&nbsp;</p>

	<div class="error">
		<img src="<?php echo CIB_PATH; ?>img/error.png" width="16" height="16" border="0" alt="" style="vertical-align:text-bottom" />
		<b><?php echo "error"; ?>:</b><br /><br />
        <?php
        echo $error;

		if ($cib_settings['debug_mode'])
		{
			echo '
            <p>&nbsp;</p>
            <p><span style="color:red;font-weight:bold">WARNING</span><br />Debug mode is enabled. Make sure you disable debug mode in settings once sONAr is installed and working properly.</p>';
		}
        ?>
	</div>
    <br />

<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>

<?php
exit();
} // END cib_error()

}

ini_set('max_execution_time', '0');
ini_set('memory_limit', '2048M');

$sql = "SELECT `mqt_rep_request`.*, `mailcronrun`.`mqt_rep`,`mailcronrun`.`bload`,`customer`.`ref` as refid,`customer`.`name` as cuname,`country`.`name` as `cname` from `mqt_rep_request`,`mailcronrun`,`customer`,`country` WHERE `run`='0' 
AND `customer`.`id`=`mqt_rep_request`.`cuid` AND `customer`.`coid`=`country`.`id` ORDER BY `when_req` ASC LIMIT 1";

$res = cib_dbQuery($sql);
$runc=cib_dbFetchAssoc($res);

if (!isSet($runc['run']) || $runc['mqt_rep'] == 1 || $runc['bload'] != 0) exit;

$custfn= str_replace(array(' ', '%','&','.','/'), '_',!Empty($runc['cuname']) ? $runc['cuname'] : 'm');
$xlsfilname=str_replace(array(' ', '%','&','.','/'), '_', $runc['opnbr'])."_sONAr_MQT_".$custfn.date_format(date_create(),'_jS_M_Y_H_m_s').".xlsx";

cib_dbQuery("update `mqt_rep_request` SET `when_started`= NOW(),`run` = '2' where `id`='".cib_dbEscape($runc['id'])."' LIMIT 1");
cib_dbQuery("update `mailcronrun` SET `mqt_rep`='1' LIMIT 1");

/* del
$sql = "SELECT `id` FROM `boardinv` ORDER BY id DESC LIMIT 1";
$res = cib_dbQuery($sql);
$tmp=cib_dbFetchAssoc($res);
cib_dbQuery("ALTER TABLE `mqt_rep` AUTO_INCREMENT =".cib_dbEscape($tmp['id']+1)."");
*/

$sql_alsm_ch = "SELECT count(*) FROM `alsm_cust_profile` WHERE `CUID`= '".$runc['cuid']."'";
//echo $sql_alsm_ch;
$sql_alsm_cha = cib_dbQuery($sql_alsm_ch);
$alsm_ntaged=cib_dbResult($sql_alsm_cha);

$alsm_not_taged=0;
$HideTab=0;

if ($alsm_ntaged) 
{
$sql_alsm_ch = "SELECT count(*) FROM `alsm_cust_profile` WHERE `CUID`= '".$runc['cuid']."' ";
$sql_alsm_cha = cib_dbQuery($sql_alsm_ch);
$alsm_ntaged=cib_dbResult($sql_alsm_cha);
$alsm_not_taged = ($alsm_ntaged == 0 ? 1 : 0); 
$HideTab = ($alsm_not_taged == 1 ? 1 : 0); 
}

function cib_RepType($in)
{
	switch ($in)
	{
		case 1:
		$reptype= "NFM-P Card (Physical Equipment)";
		break;
		case 2:
		$reptype= "NFM-P Shelf (Physical Equipment)";
		break;
		case 3:
		$reptype= "NFM-T (Global RI - from CLI)";
		break;
		case 4:
		$reptype= "IR - RM PhM";
		break;	
		case 5:
		$reptype= "NFM-P Media Adapter (Physical Equipment)";
		break;	
		case 7:
		$reptype= "NFM-T Network Element (Network)";
		break;
		case 8:
		$reptype= "NFM-T (Global RI - from GUI)";
		break;
		default:
		$reptype= "HC Tool";
	}
	return $reptype;
}

function yearfm($months)
{
  $str = '';

  if(($y = round(bcdiv($months, 12))))
  {
    $str .= "$y Year".($y-1 ? 's' : null);
  }
  if(($m = round($months % 12)))
  {
    $str .= ($y ? ' ' : null)."$m Month".($m-1 ? 's' : null);
  }
  else if ($str == '') $str='None';
  return empty($str) ? false : $str;
}

$USD_EUR=0.852;
$EUR_USD=1.1725;

$tags=$runc['ntags'];
$mtcpst=$runc['priocare'];
$mtcpst_o=$mtcpst;
$mtcstd=$runc['startcare'];
$mtcstd_o=$mtcstd;



//cib_dbQuery("INSERT IGNORE INTO `mqt_rep`(`id`, `cuid`, `whoup`, `whenup`, `whomod`, `whenmod`, `node`, `sn`, `mandate`, `ctype`, `slid`, `shtype`, `soft`, `del`, `sname`, `pn`, `tag`, `reptype`, `ftxttag`, `outage`, `eolb`) SELECT `id`, `cuid`, `whoup`, `whenup`, `whomod`, `whenmod`, `node`, `sn`, `mandate`, `ctype`, `slid`, `shtype`, `soft`, `del`, `sname`, `pn`, `tag`, `reptype`, `ftxttag`, `outage`, `eolb` FROM `boardinv` WHERE `tag` iN ($tags) AND `reptype` IN (1,2,3,5,7,8,9,11)");
cib_dbQuery("INSERT IGNORE INTO `mqt_rep`(`id`, `cuid`, `whoup`, `whenup`, `whomod`, `whenmod`, `node`, `sn`, `mandate`, `ctype`, `slid`, `shtype`, `soft`, `del`, `sname`, `pn`, `tag`, `reptype`, `ftxttag`, `outage`, `eolb`) SELECT `id`, `cuid`, `whoup`, `whenup`, `whomod`, `whenmod`, `node`, `sn`, `mandate`, `ctype`, `slid`, `shtype`, `soft`, `del`, `sname`, `pn`, `tag`, `reptype`, `ftxttag`, `outage`, `eolb` FROM `boardinv` WHERE `tag` iN ($tags) AND `reptype` IN (1,2,3,5,7,8,9,11) AND `sn` NOT IN ('N/A','')");
/* Start collecting Node summ

$sql = "SELECT DISTINCT IF( b.shtype = '7705-SAR8 v2', '7705-SAR8', b.shtype) as shtype,SUBSTRING_INDEX(b.`soft`,'R',1) as soft,
b.`tag`,c.pr_com_name,c.mqt_config,c.c_pr_name,c.mqt_product,c.mqt_u_name,t.tag as tagn
FROM `mqt_rep` AS b, shelf_mqt_conv as c,ntag as t 
WHERE b.`shtype`=c.`shelf_type_disp` AND t.id=b.`tag` AND `t`.`mngt`='1' AND `c`.`mngt`='1' AND `c`.excep='0'
ORDER BY `t`.`tag` ASC, `b`.`node` ASC";

*/
// Stopt collecing Node summ

$sql = "SELECT DISTINCT `node`,`mandate` FROM `mqt_rep` WHERE `reptype`='2' ORDER BY `mqt_rep`.`node` DESC";	
$sqlf = cib_dbQuery($sql);


	while ($tmp=cib_dbFetchAssoc($sqlf))
	{
	cib_dbQuery("update `mqt_rep` SET `mandate`= '".cib_dbEscape($tmp['mandate'])."',`fmd`='1' where `node`='".cib_dbEscape($tmp['node'])."' AND `reptype`='5' AND `mandate`='0000-00-00'");
	}	

$sql = "SELECT DISTINCT b.`cuid`,b.`node`,b.`shtype`,b.`soft`,b.`tag`,b.`reptype`,c.`IBR` 
FROM `mqt_rep` AS b, shelf_mqt_conv as c WHERE 
`shtype` IN (SELECT `shelf_type_disp` FROM `shelf_mqt_conv` WHERE `IBR` !='') AND b.`shtype`=c.`shelf_type_disp` ORDER BY `b`.`node` ASC";

$sqlibr = cib_dbQuery($sql);	

	while ($tmp=cib_dbFetchAssoc($sqlibr))
	{

$custID = $tmp['cuid'];
$SiteName= $tmp['node'];
$SerialNumber = 'IBR-'.hash( 'crc32',mt_rand(),false);
$SlotID =0;
$CardType = 'Node';
$PartNumber = $tmp['IBR'];
$Equipped = 'N/A';
$SoftwareVersion = $tmp['soft'];
$ShelfType = $tmp['shtype'];
$tag = $tmp['tag'];
$ManufactureDate = '';
$SlotName = '-';	
$ftxttag = '';
$reptype = $tmp['reptype'];
	
			$sql = "INSERT INTO `mqt_rep` (`id`,`cuid`,`whoup` ,`whenup`,`whomod`,`whenmod` ,`node` ,
			`sn` ,`mandate` ,`ctype` ,`slid` ,`shtype`,`sname`,`pn`,`tag`,`ftxttag`,`reptype`,`soft` ,`del` )
			VALUES (
			NULL ,
			'".cib_dbEscape($custID)."',
			'1',
			'',
			'' ,
			'' ,
			'".cib_dbEscape($SiteName)."',
			'".cib_dbEscape($SerialNumber)."',
			'".cib_dbEscape($ManufactureDate)."',	
			'".cib_dbEscape($CardType)."',		
			'".cib_dbEscape($SlotID)."',	
			'".cib_dbEscape($ShelfType)."',
			'".cib_dbEscape($SlotName)."',
			'".cib_dbEscape($PartNumber)."',
			'".cib_dbEscape($tag)."',
			'".cib_dbEscape($ftxttag)."',
			'".cib_dbEscape($reptype)."',
			'".cib_dbEscape($SoftwareVersion)."', '0')";
			//echo $sql."<br>";
			cib_dbQuery($sql);
	}
	
$mtcpst=date_format(date_create($mtcpst),'jS M Y');
$mtcstd=date_format(date_create($mtcstd),'jS M Y');

$mtcstd_1= strtotime($mtcstd);
$mtcstd_1 = strtotime('+ 1 year', $mtcstd_1);
$mtcstd_1 = date('Y-m-d', $mtcstd_1);
$mtcstd_o1=$mtcstd_1;
$mtcstd_1 = date_format(date_create($mtcstd_1),'jS M Y');

$mtcstd_2= strtotime($mtcstd);
$mtcstd_2 = strtotime('+ 2 year', $mtcstd_2);
$mtcstd_2 = date('Y-m-d', $mtcstd_2);
$mtcstd_o2=$mtcstd_2;
$mtcstd_2 = date_format(date_create($mtcstd_2),'jS M Y');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');
/** Include PHPSpreadsheet */
require_once dirname(__FILE__) . '/inc/Classes/PhpSpreadsheet/Spreadsheet.php';
require_once dirname(__FILE__) . '/inc/Classes/PhpSpreadsheet/Cell/AdvancedValueBinder.php';

$objPHPExcel = new PHPExcel();
$objPHPExcel->setActiveSheetIndex(0)->setTitle("Notes and Errors")->getTabColor()->setRGB('FF0000');
$objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(85);
PHPExcel_Cell::setValueBinder( new PHPExcel_Cell_AdvancedValueBinder() );
$objPHPExcel->getActiveSheet()->getProtection()->setPassword('kuku');
$objPHPExcel->getActiveSheet()->getProtection()->setSheet(false);
$objPHPExcel->getActiveSheet()->getStyle('A1:B1')->getProtection()->setLocked(PHPExcel_Style_Protection::PROTECTION_UNPROTECTED);
$objPHPExcel->getDefaultStyle()->getFont()->setName('Calibri')->setSize(11);
$objPHPExcel->getSecurity()->setLockWindows(true);
$objPHPExcel->getSecurity()->setLockStructure(true);
$objPHPExcel->getSecurity()->setWorkbookPassword('secret');
//$objPHPExcel->getDefaultStyle()->getAlignment()->setWrapText(true);

$styleArray0 = array(
	'font' => array(
		'bold' => false,
		'name'  => 'Calibri'
		//'color'		=> array('argb' => 'F4F4F4')
	),
	'alignment' => array(
		'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
		'vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,
		'wrap' => true,
	),
	'borders' => array(
		'allborders' => array(
		'style' => PHPExcel_Style_Border::BORDER_THIN,
		),
	),
	'fill' => array(
		'type' => PHPExcel_Style_Fill::FILL_SOLID,
		'color'		=> array('argb' => 'D9E1F2')
	),
);

$styleArray00 = array(
	'font' => array(
		'bold' => false,
		'name'  => 'Calibri'
		//'color'		=> array('argb' => 'F4F4F4')
	),
	'alignment' => array(
		'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
		'vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,
		'wrap' => false,
	),
	'borders' => array(
		'allborders' => array(
		'style' => PHPExcel_Style_Border::BORDER_THIN,
		),
	),
	'fill' => array(
		'type' => PHPExcel_Style_Fill::FILL_SOLID,
		'color'		=> array('argb' => 'DDD9C4')
	),
);

$styleArray01 = array(
	'font' => array(
		'bold' => false,
		'name'  => 'Calibri'
		//'color'		=> array('argb' => 'F4F4F4')
	),
	'alignment' => array(
		'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
		'vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,
		'wrap' => false,
	),
	'borders' => array(
		'allborders' => array(
		'style' => PHPExcel_Style_Border::BORDER_THIN,
		),
	),
	'fill' => array(
		'type' => PHPExcel_Style_Fill::FILL_SOLID,
		//'color'		=> array('argb' => 'D9E1F2')
	),
);

$styleArray01_L = array(
	'font' => array(
		'bold' => false,
		'name'  => 'Calibri',
		'color'	=> array('argb' => '0000FF'),
		'underline' => 'single'
	),
	'alignment' => array(
		'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
		'vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,
		'wrap' => false,
	),
	'borders' => array(
		'allborders' => array(
		'style' => PHPExcel_Style_Border::BORDER_THIN,
		),
	),
	'fill' => array(
		'type' => PHPExcel_Style_Fill::FILL_SOLID,
		//'color'		=> array('argb' => 'D9E1F2')
	),
);

$styleArray01G = array(
	'font' => array(
		'bold' => false,
		'name'  => 'Calibri'
		//'color'		=> array('argb' => 'F4F4F4')
	),
	'alignment' => array(
		'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
		'vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,
		'wrap' => true,
	),
	'borders' => array(
		'allborders' => array(
		'style' => PHPExcel_Style_Border::BORDER_THIN,
		),
	),
	'fill' => array(
		'type' => PHPExcel_Style_Fill::FILL_SOLID,
		'color'		=> array('argb' => 'D8E4BC')
	),
);

$styleArray02 = array(
	'font' => array(
		'bold' => false,
		'name'  => 'Calibri'
		//'color'		=> array('argb' => 'F4F4F4')
	),
	'alignment' => array(
		'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
		'vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,
		'wrap' => true,
	),
	'borders' => array(
		'allborders' => array(
		'style' => PHPExcel_Style_Border::BORDER_THIN,
		),
	),
	'fill' => array(
		'type' => PHPExcel_Style_Fill::FILL_SOLID,
		'color'		=> array('argb' => 'FCD5B4')
	),
);

$styleArray03 = array(
	'font' => array(
		'bold' => false,
		'name'  => 'Calibri'
		//'color'		=> array('argb' => 'F4F4F4')
	),
	'alignment' => array(
		'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
		'vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,
		'wrap' => true,
	),
	'borders' => array(
		'allborders' => array(
		//'style' => PHPExcel_Style_Border::BORDER_THIN,
		),
	),
	'fill' => array(
		'type' => PHPExcel_Style_Fill::FILL_SOLID,
		'color'		=> array('argb' => '7B7B7B')
	),
);

$styleArray1 = array(
	'font' => array(
		'bold' => True,
		'name'  => 'Calibri'
		//'color'		=> array('argb' => 'F4F4F4')
	),
	'alignment' => array(
		'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
		'vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,
		'wrap' => true,
	),
	'borders' => array(
		'allborders' => array(
		'style' => PHPExcel_Style_Border::BORDER_THIN,
		),
	),
	'fill' => array(
		'type' => PHPExcel_Style_Fill::FILL_SOLID,
		'color'		=> array('argb' => 'D9E1F2')
	),
);

$objPHPExcel->getProperties()->setCreator('sONAr')
							 ->setLastModifiedBy('sONAr')
							 ->setTitle('Test')
							 ->setSubject('Test')
							 ->setDescription('Test')
							 ->setKeywords('Test')
							 ->setCategory('Test');
							 
$i=8;
$iA=8;
$cnt=1;
$cnt_ns=1;
$i_ns=8;
$redA=0;
$sql6 = "SELECT distinct mqt_rep.tag as tag,ntag.tag as tagn FROM `mqt_rep`,`ntag` WHERE `mqt_rep`.tag=ntag.id AND ntag.mngt='1' AND `mqt_rep`.`shtype` NOT IN ('N/A') ORDER by tag";	
$res16 = cib_dbQuery($sql6);
$is_el_tr=cib_dbNumRows($res16);

$sql66 = "SELECT distinct alsm_cust_profile.NTAG as tag,ntag.tag as tagn FROM `alsm_cust_profile`,`ntag` WHERE `alsm_cust_profile`.NTAG=ntag.id AND `NTAG` IN ($tags)";	
$res166 = cib_dbQuery($sql66);
$is_alsm_ok=cib_dbNumRows($res166);

$sql_ntag = "SELECT * FROM `ntag` WHERE `id` IN ($tags)";
$ntags=cib_dbQuery($sql_ntag);
$ntags_f='';

while ($ntag_a=cib_dbFetchAssoc($ntags))
{
	$ntags_f.=$ntag_a['tag'].", ";
}
$ntags_f=substr($ntags_f,0,-2);

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth('30');
$objPHPExcel->getActiveSheet()->setCellValue('A1', "Care Renewal Assessment (CRA)")->getStyle('A3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('A2', "Opportunity number:")->getStyle('A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$objPHPExcel->getActiveSheet()->setCellValue('C2', "Deepfield, and Video products are not supported by CRA")->getStyle('C2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('C2')->getFont()->getColor()->setRGB('FF0000');
$objPHPExcel->getActiveSheet()->getStyle('C2')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->mergeCells('C2:E2');

$objPHPExcel->getActiveSheet()->setCellValue('A3', "Date Report Generated:")->getStyle('A3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Date Report Generated:")->getStyle('A3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->setCellValue('A4', "Customer Name:")->getStyle('A4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->setCellValue('A5', "Country:")->getStyle('A5')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->setCellValue('A6', "Customer Ref. Nb.:")->getStyle('A6')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->setCellValue('A7', "NTAG(s) included:")->getStyle('A7')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->setCellValue('A8', "Currency:")->getStyle('A8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->setCellValue('A9', "Prior Care Contract start date:")->getStyle('A9')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('A9')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('A10', "New Care Contract start date:")->getStyle('A10')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('A10')->getFont()->getColor()->setRGB('0070C0');
//$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth('40');
$objPHPExcel->getActiveSheet()->setCellValue('B2', $runc['opnbr'])->getStyle('B2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('B3', date('jS M Y'))->getStyle('B3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('B4', $runc['cuname'])->getStyle('B4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('B5', $runc['cname'])->getStyle('B5')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('B6', $runc['refid'])->getStyle('B6')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('B7', $ntags_f)->getStyle('B7')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('B8', $runc['curr'])->getStyle('B8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('B9', $mtcpst)->getStyle('B9')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('B9')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('B10', $mtcstd)->getStyle('B10')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('B10')->getFont()->getColor()->setRGB('0070C0');

$objPHPExcel->getActiveSheet()->setCellValue('A13', "Electronic Audit Files from NMS Sources:")->getStyle('A13')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('A14')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('B14', "NTAG")->getStyle('B14')->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getStyle('B14')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('C14', "NMS Source")->getStyle('C14')->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getStyle('C14')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('D14', "File")->getStyle('D14')->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getStyle('D14')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('E14', "File Contents")->getStyle('E14')->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getStyle('E14')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('F14', "Upload date")->getStyle('F14')->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getStyle('F14')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('G14', "Errors")->getStyle('G14')->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getStyle('G14')->getFont()->setBold( true );

$i1=15;

$RtypesA = [
			'1' => 'NFM-P:Card (Physical Equipment):1 of 4',
			'2' => 'NFM-P:Shelf (Physical Equipment):2 of 4',
			'5' => 'NFM-P:Media Adapter (Physical Equipment):3 of 4',
			'7' => 'NFM-P:Network Element (Network):4 of 4',
			'3' => 'NFM-T:Global RI & NE List - from CLI:1&2 of 2',	
			'8' => 'NFM-T:Global RI & NE List - from GUI:1&2 of 2',
			'9' => 'HC Tool:HC Tool:2 of 2'
			];

function dateDiff($date1, $date2)  //days find function
	{ 
        $diff = strtotime($date2) - strtotime($date1); 
        return abs(round($diff / 86400)); 
    } 

$sql = "SELECT DISTINCT `ntag`.`id`,`ntag`.`tag` FROM `mqt_rep`,`ntag` where `ntag`.`id`=`mqt_rep`.`tag` ORDER BY `ntag`.`tag`,`reptype` ASC";
$sqlr = cib_dbQuery($sql);
	while ($tmp=cib_dbFetchAssoc($sqlr))
	{	
		$isIP=0;
		$DateRef='0000-00-00 00:00:00';
		$filetypea=Array();
		foreach ($RtypesA as $key => $value) 
		{
		$sql1 = "SELECT `reptype`,if (max(`whenmod`)='0000-00-00 00:00:00',max(`whenup`),max(`whenmod`)) as `whenup` FROM `mqt_rep` where `mqt_rep`.`tag`='".cib_dbEscape($tmp['id'])."' AND `reptype`='".$key."' ";
		$sqlr1 = cib_dbQuery($sql1);
		$repa=cib_dbFetchAssoc($sqlr1);
		$filetypea=explode(':',$value);
		$fdatetxt='';
		If ($repa['reptype']==$key)
			{		
				if ($key==1) {$isIP=1;$DateRef=$repa['whenup'];}			
				if ($isIP AND in_array($key, array(1,2,5,7))) $diff=dateDiff($DateRef,$repa['whenup']);	
				else $diff=0;
				
				$sqlf = "SELECT COUNT(*) FROM `mqt_rep` WHERE `mandate` > Now() AND `reptype`='$key' ";
				$sqlfa = cib_dbQuery($sqlf);
				$fdate=cib_dbResult($sqlfa);
				//echo $fdate.'<br>';
				if ($fdate>0) {
				$fdatetxt="\rNumber of units with manufacturing date in the future detected -".$fdate;
				$redA=1;
				}
				$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, $tmp['tag'])->getStyle('B'.$i1)->applyFromArray($styleArray01);
				$objPHPExcel->getActiveSheet()->setCellValue('C'.$i1, $filetypea[0])->getStyle('C'.$i1)->applyFromArray($styleArray01);
				$objPHPExcel->getActiveSheet()->setCellValue('D'.$i1, $filetypea[2])->getStyle('D'.$i1)->applyFromArray($styleArray01);
				$objPHPExcel->getActiveSheet()->setCellValue('E'.$i1, $filetypea[1])->getStyle('E'.$i1)->applyFromArray($styleArray01);
				$objPHPExcel->getActiveSheet()->setCellValue('F'.$i1, date_format(date_create($repa['whenup']),'jS F Y'))->getStyle('F'.$i1)->applyFromArray($styleArray01);
				$objPHPExcel->getActiveSheet()->setCellValue('G'.$i1, $fdatetxt)->getStyle('G'.$i1)->applyFromArray($styleArray01);
				$objPHPExcel->getActiveSheet()->getStyle('G'.$i1)->getFont()->getColor()->setRGB('FF0000');
				
				if ($diff > 30)
				{
				$redA=1;
				$HideTab=1;
				$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFont()->setBold(true)->getColor()->setRGB('FF0000');
				$objPHPExcel->getActiveSheet()->getStyle('C'.$i1)->getFont()->setBold(true)->getColor()->setRGB('FF0000');
				$objPHPExcel->getActiveSheet()->getStyle('D'.$i1)->getFont()->setBold(true)->getColor()->setRGB('FF0000');
				$objPHPExcel->getActiveSheet()->getStyle('E'.$i1)->getFont()->setBold(true)->getColor()->setRGB('FF0000');
				$objPHPExcel->getActiveSheet()->getStyle('F'.$i1)->getFont()->setBold(true)->getColor()->setRGB('FF0000');
				$objPHPExcel->getActiveSheet()->setCellValue('G'.$i1, "File age is ".$diff." days different than the Card Report. Consider reuploading it.".$fdatetxt)->getStyle('G'.$i1)->applyFromArray($styleArray01);
				$objPHPExcel->getActiveSheet()->getStyle('G'.$i1)->getAlignment()->setWrapText(true);
				$objPHPExcel->getActiveSheet()->getStyle('G'.$i1)->getFont()->getColor()->setRGB('FF0000');
				}
				$i1++;
			}
		else if ($isIP AND in_array($key, array(2,5,7)))
			{
				$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, $tmp['tag'])->getStyle('B'.$i1)->applyFromArray($styleArray01);
				$objPHPExcel->getActiveSheet()->setCellValue('C'.$i1, $filetypea[0])->getStyle('C'.$i1)->applyFromArray($styleArray01);
				$objPHPExcel->getActiveSheet()->setCellValue('D'.$i1, $filetypea[2])->getStyle('D'.$i1)->applyFromArray($styleArray01);
				$objPHPExcel->getActiveSheet()->setCellValue('E'.$i1, $filetypea[1])->getStyle('E'.$i1)->applyFromArray($styleArray01);
				$objPHPExcel->getActiveSheet()->setCellValue('F'.$i1, 'Not Uploaded')->getStyle('F'.$i1)->applyFromArray($styleArray01);
				$objPHPExcel->getActiveSheet()->setCellValue('G'.$i1, $fdatetxt)->getStyle('G'.$i1)->applyFromArray($styleArray01);
				$objPHPExcel->getActiveSheet()->getStyle('G'.$i1)->getFont()->getColor()->setRGB('FF0000');
				$i1++;
			}				

		}
	}

	$i1++;
	
	$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, "Table A ")->getStyle('B'.$i1)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFont()->setBold( true );
	
	$i1++;
	$il2=$i1;
	
	$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, "NTAG")->getStyle('B'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFont()->setBold( true );	
	$objPHPExcel->getActiveSheet()->setCellValue('C'.$i1, "Product")->getStyle('C'.$i1)->applyFromArray($styleArray01);	
	$objPHPExcel->getActiveSheet()->getStyle('C'.$i1)->getFont()->setBold( true );
	$objPHPExcel->getActiveSheet()->setCellValue('D'.$i1, "Release")->getStyle('D'.$i1)->applyFromArray($styleArray01);	
	$objPHPExcel->getActiveSheet()->getStyle('D'.$i1)->getFont()->setBold( true );
	
	$i1++;
	$reset=1;
	$sqlr = cib_dbQuery($sql);	
	while ($tmp=cib_dbFetchAssoc($sqlr))
	{
		$sql1 = "SELECT `dpr`.prname AS `dpproduct`,`dr`.relase AS `drrelease`,	ntag.tag AS tag, `dpr`.id as prid
		FROM `itb` AS `cib`
		LEFT JOIN `prrelease` AS `dr` on `cib`.`dprrelid`=`dr`.`id`
		LEFT JOIN `product` AS `dpr` on `dr`.`prid`=`dpr`.`id`
		LEFT JOIN `prfam` AS `pf` on `dpr`.`prfid`=`pf`.`id`
		LEFT JOIN `techno` AS `t` on `pf`.`techid`=`t`.`id`
		LEFT JOIN `ntag` AS `ntag` on `cib`.`ntag`=`ntag`.`id`
		LEFT JOIN `shelf_mqt_conv` AS mqt on `mqt`.`s_name_id`=`dpr`.id
		WHERE `cib`.`del`='0' AND `cib`.`ntag` ='".cib_dbEscape($tmp['id'])."' 
		AND (`mqt`.`mngt`='1' AND `mqt`.excep='0')
		AND `dpr`.id NOT IN (SELECT DISTINCT `conv`.`s_name_id` FROM `mqt_rep` as `rep`,`shelf_mqt_conv` as `conv` WHERE `conv`.`shelf_type_disp`=`rep`.`shtype` AND `tag` ='".cib_dbEscape($tmp['id'])."') GROUP BY `cib`.`id`";
		$sqlr1 = cib_dbQuery($sql1);
		while ($tmp1=cib_dbFetchAssoc($sqlr1))
		{	
		$reset=0;
		$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, $tmp1['tag'])->getStyle('B'.$i1)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('C'.$i1, $tmp1['dpproduct'])->getStyle('C'.$i1)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('D'.$i1, '\''.$tmp1['drrelease'])->getStyle('D'.$i1)->applyFromArray($styleArray01);
		$i1++;
		}	
		

	}

		if ($reset) {
		$objPHPExcel->getActiveSheet()->setCellValue('B'.($il2-1), '')->getStyle('B'.($il2-1))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
		$objPHPExcel->getActiveSheet()->setCellValue('B'.$il2, '')->getStyle('B'.$il2)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_HAIR);
		$objPHPExcel->getActiveSheet()->setCellValue('C'.$il2, '')->getStyle('C'.$il2)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_HAIR);
		$objPHPExcel->getActiveSheet()->setCellValue('D'.$il2, '')->getStyle('D'.$il2)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_HAIR);
		$i1=$il2-2;
		}

if (!$reset) {
		
$i1++;

	$objPHPExcel->getActiveSheet()->setCellValue('E'.$i1, "Next Steps")->getStyle('E'.$i1)->applyFromArray($styleArray01);	
	$objPHPExcel->getActiveSheet()->getStyle('E'.$i1)->getFont()->setBold( true );
	$objPHPExcel->getActiveSheet()->setCellValue('F'.$i1, "Link 1")->getStyle('F'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('F'.$i1)->getFont()->setBold( true );
	$objPHPExcel->getActiveSheet()->setCellValue('G'.$i1, "Link 2")->getStyle('G'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('G'.$i1)->getFont()->setBold( true );

$i1++;
	$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, "See Table A for products found in the Manual Track (\"IB Nodal Summary\") that can be audited electronically by the attached NMS systems. (Note that software releases are ignored by this check)")->getStyle('B'.$i1)->applyFromArray($styleArray01);	
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->setCellValue('C'.$i1, "")->getStyle('C'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('C'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->setCellValue('D'.$i1, "")->getStyle('D'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('D'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->mergeCells('B'.$i1.':'.'D'.$i1);
	$objPHPExcel->getActiveSheet()->getRowDimension($i1)->setRowHeight(45);
	$objPHPExcel->getActiveSheet()->setCellValue('E'.$i1, "Run the NMS audit and upload into Sonar.")->getStyle('E'.$i1)->applyFromArray($styleArray01);	
	$objPHPExcel->getActiveSheet()->setCellValue('F'.$i1, "How to")->getStyle('F'.$i1) ->applyFromArray($styleArray01_L);
	//$objPHPExcel->getActiveSheet()->setCellValue('F'.$i1, '=Hyperlink("https://nokia.sharepoint.com/sites/sONAr/blog/Lists/Posts/Post.aspx?ID=67","How to")')->getStyle('F'.$i1)->applyFromArray($styleArray01);
	$url="https://nokia.sharepoint.com/sites/sONAr/blog/Lists/Posts/Post.aspx?ID=67";
	$objPHPExcel->getActiveSheet()->getHyperlink('F'.$i1)->setUrl($url);	
	$objPHPExcel->getActiveSheet()->setCellValue('G'.$i1, "")->getStyle('G'.$i1)->applyFromArray($styleArray01);
$i1++;
	$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, "Note:  If you believe you are receiving this message in error (because you did upload an audit)  there could be a problem with:")->getStyle('B'.$i1)->applyFromArray($styleArray01);	
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');
	$objPHPExcel->getActiveSheet()->setCellValue('C'.$i1, "")->getStyle('C'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('C'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->setCellValue('D'.$i1, "")->getStyle('D'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('D'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->mergeCells('B'.$i1.':'.'D'.$i1);
	$objPHPExcel->getActiveSheet()->getRowDimension($i1)->setRowHeight(45);
	$objPHPExcel->getActiveSheet()->setCellValue('E'.$i1, "")->getStyle('E'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('E'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');	
	$objPHPExcel->getActiveSheet()->setCellValue('F'.$i1, "")->getStyle('F'.$i1) ->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('F'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');	
	$objPHPExcel->getActiveSheet()->setCellValue('G'.$i1, "")->getStyle('G'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('G'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');	
$i1++;
	$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, "a) gap or error in internal sONAr datatables OR")->getStyle('B'.$i1)->applyFromArray($styleArray01);	
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');
	$objPHPExcel->getActiveSheet()->setCellValue('C'.$i1, "")->getStyle('C'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('C'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->setCellValue('D'.$i1, "")->getStyle('D'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('D'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->mergeCells('B'.$i1.':'.'D'.$i1);
	$objPHPExcel->getActiveSheet()->getRowDimension($i1)->setRowHeight(45);
	$objPHPExcel->getActiveSheet()->setCellValue('E'.$i1, "Contact the Sonar Admin or the Care PLM\nresponsible for the product")->getStyle('E'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('E'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->getStyle('E'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');	
	$objPHPExcel->getActiveSheet()->setCellValue('F'.$i1, "Sonar Admin")->getStyle('F'.$i1) ->applyFromArray($styleArray01_L);
	$url="http://ddtb.de.alcatel-lucent.com/cib/support.php";
	$objPHPExcel->getActiveSheet()->getHyperlink('F'.$i1)->setUrl($url);	
	$objPHPExcel->getActiveSheet()->getStyle('F'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');	
	$objPHPExcel->getActiveSheet()->setCellValue('G'.$i1, "Care PLM Prime")->getStyle('G'.$i1)->applyFromArray($styleArray01_L);
	$objPHPExcel->getActiveSheet()->getStyle('G'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');
	$url="http://nok.it/CarePLMPrime";
	$objPHPExcel->getActiveSheet()->getHyperlink('G'.$i1)->setUrl($url);	
$i1++;
	$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, "b) the IB Nodal Summary (Manual Track) is incorrect OR")->getStyle('B'.$i1)->applyFromArray($styleArray01);	
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');
	$objPHPExcel->getActiveSheet()->setCellValue('C'.$i1, "")->getStyle('C'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('C'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->setCellValue('D'.$i1, "")->getStyle('D'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('D'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->mergeCells('B'.$i1.':'.'D'.$i1);
	$objPHPExcel->getActiveSheet()->getRowDimension($i1)->setRowHeight(45);
	$objPHPExcel->getActiveSheet()->setCellValue('E'.$i1, "Go to the IB Nodal Summary and delete\nthe products that are not present in\nthe network")->getStyle('E'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('E'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->getStyle('E'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');	
	$objPHPExcel->getActiveSheet()->setCellValue('F'.$i1, "Nodal Summary page")->getStyle('F'.$i1) ->applyFromArray($styleArray01_L);
	//$url="http://ddtb.de.alcatel-lucent.com/cib/support.php";
	$url=$cib_settings['site_url'].'acib.php?custID='.$runc['cuid'];
	$objPHPExcel->getActiveSheet()->getHyperlink('F'.$i1)->setUrl($url);	
	$objPHPExcel->getActiveSheet()->getStyle('F'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');	
	$objPHPExcel->getActiveSheet()->setCellValue('G'.$i1, "")->getStyle('G'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('G'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');
$i1++;
	$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, "c) the upload itself, because it was corrupted")->getStyle('B'.$i1)->applyFromArray($styleArray01);	
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');
	$objPHPExcel->getActiveSheet()->setCellValue('C'.$i1, "")->getStyle('C'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('C'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->setCellValue('D'.$i1, "")->getStyle('D'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('D'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->mergeCells('B'.$i1.':'.'D'.$i1);
	$objPHPExcel->getActiveSheet()->getRowDimension($i1)->setRowHeight(45);
	$objPHPExcel->getActiveSheet()->setCellValue('E'.$i1, "Reupload the NMS audit")->getStyle('E'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('E'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->getStyle('E'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');	
	$objPHPExcel->getActiveSheet()->setCellValue('F'.$i1, "Audit upload page")->getStyle('F'.$i1) ->applyFromArray($styleArray01_L);
	$url=$cib_settings['site_url'].'binv.php?custID='.$runc['cuid'];
	$objPHPExcel->getActiveSheet()->getHyperlink('F'.$i1)->setUrl($url);	
	$objPHPExcel->getActiveSheet()->getStyle('F'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');	
	$objPHPExcel->getActiveSheet()->setCellValue('G'.$i1, "Sonar Admin")->getStyle('G'.$i1)->applyFromArray($styleArray01_L);
	$objPHPExcel->getActiveSheet()->getStyle('G'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');
	$url="http://nok.it/CarePLMPrime";
	$objPHPExcel->getActiveSheet()->getHyperlink('G'.$i1)->setUrl($url);	
$i1++;
}
	
$i1++;

$sql = "SELECT DISTINCT `shtype` FROM `mqt_rep` WHERE `shtype` NOT IN (SELECT `shelf_type_disp` FROM `shelf_mqt_conv` WHERE 1)";
$sqlr = cib_dbQuery($sql);
$not_supp_sh=cib_dbNumRows($sqlr);

if ($not_supp_sh)
{
	//$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, "Not supported shelfs detected")->getStyle('B'.$i1)->applyFromArray($styleArray01);	
	$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, "Table B ")->getStyle('B'.$i1)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFont()->setBold( true );
	$i1++;
	$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, "Shelf Type")->getStyle('B'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFont()->setBold( true );
	$objPHPExcel->getActiveSheet()->setCellValue('C'.$i1, "Amount of records")->getStyle('C'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('C'.$i1)->getFont()->setBold( true );
	
	while ($tmp=cib_dbFetchAssoc($sqlr))
	{
		$i1++;
		
		$sql_ns = "SELECT count(*) FROM `mqt_rep` WHERE `shtype`='".$tmp['shtype']."'";
		$nsamount = cib_dbResult(cib_dbQuery($sql_ns));
		$redA=1;		
		
		//echo $tmp['shtype'].'<BR>';
		$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, $tmp['shtype'])->getStyle('B'.$i1)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFont()->setBold(true)->getColor()->setRGB('FF0000');
		$objPHPExcel->getActiveSheet()->setCellValue('C'.$i1, $nsamount)->getStyle('C'.$i1)->applyFromArray($styleArray01);
	}
$i1++;
$i1++;
	$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, "This CRA is incomplete,\nunsupported shelf type\ndetected. Data was\ndetected in table B that\ncould not be\nprocessed.")->getStyle('B'.$i1)->applyFromArray($styleArray01);	
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFont()->setBold( true );
	//$objPHPExcel->getActiveSheet()->getRowDimension($i1)->setRowHeight(45);
	$objPHPExcel->getActiveSheet()->setCellValue('C'.$i1, "Contact the Sonar\nAdmin or the Care PLM\nresponsible for the\nproduct")->getStyle('C'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('C'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->getStyle('C'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');	
	$objPHPExcel->getActiveSheet()->setCellValue('D'.$i1, "Sonar Admin")->getStyle('D'.$i1) ->applyFromArray($styleArray01_L);
	$url="http://ddtb.de.alcatel-lucent.com/cib/support.php";
	$objPHPExcel->getActiveSheet()->getHyperlink('D'.$i1)->setUrl($url);	
	$objPHPExcel->getActiveSheet()->getStyle('D'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');	
	$objPHPExcel->getActiveSheet()->setCellValue('E'.$i1, "Care PLM Prime")->getStyle('E'.$i1)->applyFromArray($styleArray01_L);
	$objPHPExcel->getActiveSheet()->getStyle('E'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');
	$url="http://nok.it/CarePLMPrime";
	$objPHPExcel->getActiveSheet()->getHyperlink('E'.$i1)->setUrl($url);	
}
$cnt_mts=$i1+2;

/*
If ($redA==1){
$objPHPExcel->getActiveSheet()->setCellValue('A12', "This CRA is INCOMPLETE or contains ERRORS and should not be used.  Please take appropriate Next Steps to correct the issue and re-run the CRA.")->getStyle('A12')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('A12')->getFont()->getColor()->setRGB('FF0000');
$objPHPExcel->getActiveSheet()->getStyle('A12')->getFont()->setBold( true );
}
*/
$tabsnames=Array();
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(1)->setTitle("Summary")->getTabColor()->setRGB('FF0000');
$tabsnames[]="Summary";
$objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(85);
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(2)->setTitle("Node Summary")->getTabColor()->setRGB('FF0000');
$tabsnames[]="Node Summary";
$objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(85);

//Start Node summary header start
$objPHPExcel->getActiveSheet()->setCellValue('C1', "Node Summary")->getStyle('C1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('C1')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->getStyle('C1')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('C2', "This is the summary of all nodes separated by NTAG")->getStyle('C2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('C2')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->getStyle('C2')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('I1', "**** Changes made to this tab will NOT be reflected in the Summary tab ****")->getStyle('I1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('I1')->getFont()->setSize(12)->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->getStyle('I1')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->mergeCells('I1:K1');


$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('B7', "Item")->getStyle('B7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('C7', 'Sonar Audit Source')->getStyle('C7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('D7', 'NTAG')->getStyle('D7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('E7', "Product Name")->getStyle('E7')->applyFromArray($styleArray0);	
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('F7', 'Shelf type (name in NMS)')->getStyle('F7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('G7', 'Product "marketing name" 2')->getStyle('G7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('H7', 'MQT Product Name')->getStyle('H7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('I7', 'MQT Product Config')->getStyle('I7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('J7', 'SW Release')->getStyle('J7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('K7', 'SW Release 2')->getStyle('K7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('L7', 'Qty of Nodes')->getStyle('L7')->applyFromArray($styleArray0);
//$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth('20');
$objPHPExcel->getActiveSheet()->setCellValue('M5', "Note: 1830 PSS product can have multiple extension shelves, column '1830 PSS Shelf count' shows total quantity of shelves (main+extension) and main shelf type. ")->getStyle('M5')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('M6', "For shelf type and quantity details, go to 'Electronic Audit Details' tab ")->getStyle('M6')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('M7', '1830 PSS Shelf count')->getStyle('M7')->applyFromArray($styleArray0);
//$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('N7', 'Qty of parts')->getStyle('N7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('O7', 'MQT Unit Qty TOTAL')->getStyle('O7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('P')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('P7', 'MQT Unit Name')->getStyle('P7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('Q7', '')->getStyle('Q7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('R7', "GLP (".$runc['curr'].") TOTAL")->getStyle('R7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('S')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('S7', 'OPD%')->getStyle('S7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('T')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('T7', "PSP ".$runc['curr'].") TOTAL")->getStyle('T7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('U')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('U7', 'PSP Confidence')->getStyle('U7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('U')->setAutoSize(true);

/*
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('N7', 'OPD%')->getStyle('N7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('O7', 'PSP Confidence')->getStyle('O7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('P')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('P7', "PSP (".$runc['curr'].") EACH - ESTIMATED")->getStyle('P7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('Q7', "PSP (".$runc['curr'].") TOTAL - ESTIMATED")->getStyle('Q7')->applyFromArray($styleArray0);
*/
//END Node summary header start



$tab_cnt=2;

If ($is_alsm_ok) {
$tab_cnt++;
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex($tab_cnt)->setTitle("ASLM License Audit Details")->getTabColor()->setRGB('FF0000');
$tabsnames[]="ASLM License Audit Details";
$objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(85);
$objPHPExcel->getActiveSheet()->setCellValue('A1', "Care Renewal Assessment (CRA)")->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('A2', "ASLM License Audit Details")->getStyle('A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('A2')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->getStyle('A2')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('A3', "This table lists the Part by part breakdown found for any products that Sonar's Audit Source is from an ASLM audit  (see Column A in Summary tab)")->getStyle('A3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('A3')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('J1', "**** Changes made to this tab will NOT be reflected in the Summary tab ****")->getStyle('J1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('J1')->getFont()->setSize(12)->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->getStyle('J1')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->mergeCells('J1:L1');
$objPHPExcel->getActiveSheet()->setCellValue('B7', "Item")->getStyle('B7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('C7', 'Sonar Audit Source')->getStyle('C7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('D7', 'NTAG')->getStyle('D7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('E7', 'Part Number')->getStyle('E7')->applyFromArray($styleArray0);	
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('F7', 'Part Description')->getStyle('F7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('G7', 'Product')->getStyle('G7')->applyFromArray($styleArray0);	
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('H7', 'Qty of parts')->getStyle('H7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('I7', '')->getStyle('I7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('J7', 'Warning')->getStyle('J7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('K7', 'MQT Product Name')->getStyle('K7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('L7', 'MQT Product Config')->getStyle('L7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('M7', 'SW Release(s)')->getStyle('M7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('N6', "1 USD = ".$USD_EUR." EUR")->getStyle('N6')->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('N7', "List Price (".$runc['curr'].") \nEACH")->getStyle('N7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('O7', "List Price (".$runc['curr'].")\nTOTAL")->getStyle('O7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('P')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('P7', "Customer\nDiscount")->getStyle('P7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('Q7', "Confidence in\ncustomer\nDiscount")->getStyle('Q7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('R7', 'PSP ('.$runc['curr'].') Each ')->getStyle('R7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('S')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('S7', "PSP (".$runc['curr'].") Total")->getStyle('S7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('T')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('T7', "MQT Unit\nEACH")->getStyle('T7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('U')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('U7', "MQT Unit\nTOTAL")->getStyle('U7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('V')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('V7', "MQT Unit Name")->getStyle('V7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('W')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('W7', "TS Warranty\nLength")->getStyle('W7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->setCellValue('X3', "Technical Support (TS)")->getStyle('V3');
$objPHPExcel->getActiveSheet()->setCellValue('AB2', "Prior Care Contract start date: $mtcpst")->getStyle('Z2');
$objPHPExcel->getActiveSheet()->setCellValue('AB3', "New Care Contract start date: $mtcstd")->getStyle('Z2');
$objPHPExcel->getActiveSheet()->freezePane('A8');
$objPHPExcel->getActiveSheet()->mergeCells('X6:Z6');
$objPHPExcel->getActiveSheet()->setCellValue('X6', "Part Number Quantities")->getStyle('X6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('Y6', "")->getStyle('Y6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('Z6', "")->getStyle('Z6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('X7', "IW")->getStyle('X7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('Y7', "OOW")->getStyle('Y7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('Z7', "Total")->getStyle('Z7')->applyFromArray($styleArray00);

$objPHPExcel->getActiveSheet()->mergeCells('AA6:AC6');
$objPHPExcel->getActiveSheet()->setCellValue('AA6', "Product Selling Price (PSP) values (".$runc['curr'].")")->getStyle('AA6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AB6', "")->getStyle('AB6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AC6', "")->getStyle('AC6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AA7', "IW")->getStyle('AA7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AB7', "OOW")->getStyle('AB7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('AC')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AC7', "Total")->getStyle('AC7')->applyFromArray($styleArray00);

$objPHPExcel->getActiveSheet()->mergeCells('AD6:AG6');
$objPHPExcel->getActiveSheet()->setCellValue('AD6', "MQT Unit Quantities")->getStyle('AD6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AE6', "")->getStyle('AE6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AF6', "")->getStyle('AF6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AG6', "")->getStyle('AG6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('AD')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AD7', "IW")->getStyle('AD7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('AE')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AE7', "OOW")->getStyle('AE7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('AF')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AF7', "Total")->getStyle('AF7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('AG')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('AG7', "MQT Unit Name")->getStyle('AG7')->applyFromArray($styleArray00);

$objPHPExcel->getActiveSheet()->mergeCells('X5:AG5');
$objPHPExcel->getActiveSheet()->setCellValue('X5', "Historical View. Deployed as of Prior Contract Date as of $mtcpst")->getStyle('X5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('Y5', "")->getStyle('Y5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('Z5', "")->getStyle('Z5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AA5', "")->getStyle('AA5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AB5', "")->getStyle('AB5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AC5', "")->getStyle('AC5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AD5', "")->getStyle('AD5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AE5', "")->getStyle('AE5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AF5', "")->getStyle('AF5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AG5', "")->getStyle('AG5')->applyFromArray($styleArray00);

$objPHPExcel->getActiveSheet()->mergeCells('AH6:AJ6');
$objPHPExcel->getActiveSheet()->setCellValue('AH6', "Part Number Quantities")->getStyle('AH6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AI6', "")->getStyle('AI6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AJ6', "")->getStyle('AJ6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AH')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AH7', "IW")->getStyle('AH7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AI')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AI7', "OOW")->getStyle('AI7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AJ')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AJ7', "Total")->getStyle('AJ7')->applyFromArray($styleArray02);

$objPHPExcel->getActiveSheet()->mergeCells('AK6:AM6');
$objPHPExcel->getActiveSheet()->setCellValue('AK6', "Product Selling Price (PSP) values (".$runc['curr'].")")->getStyle('AK6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AL6', "")->getStyle('AL6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AM6', "")->getStyle('AM6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AK')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AK7', "IW")->getStyle('AK7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AL')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AL7', "OOW")->getStyle('AL7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AM')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AM7', "Total")->getStyle('AM7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('AN6:AP6');
$objPHPExcel->getActiveSheet()->setCellValue('AN6', "MQT Unit Quantities")->getStyle('AN6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AO6', "")->getStyle('AO6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AP6', "")->getStyle('AP6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AN')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AN7', "IW")->getStyle('AN7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AO')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AO7', "OOW")->getStyle('AO7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AP')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AP7', "Total")->getStyle('AP7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('AH5:AP5');
$objPHPExcel->getActiveSheet()->setCellValue('AH5', "New Contract View: 1st year as of $mtcstd")->getStyle('AH5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AI5', "")->getStyle('AI5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AJ5', "")->getStyle('AJ5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AK5', "")->getStyle('AK5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AL5', "")->getStyle('AL5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AM5', "")->getStyle('AM5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AN5', "")->getStyle('AN5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AO5', "")->getStyle('AO5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AP5', "")->getStyle('AP5')->applyFromArray($styleArray02);

$Node_collector_alsm=Array();

	while ($ticket166=cib_dbFetchAssoc($res166))
	{
		$ntag_n=$ticket166['tag'];
		$ntag_n_d=$ticket166['tagn'];
		
		$sql55 = "SELECT `info`.`LICENSE_ID`,`info`.`ROOT_LICENSE_ID` as `rlid`,`info`.`PRODUCT_NAME` as `product`,`info`.`CUSTOMER_NAME`,`info`.`CUSTOMER_NUMBER`,`info`.`COUNTRY_NAME` 
		,`info`.`PRODUCT_RELEASE_NAME` as `release`,`info`.`CUSTOMER_SYSTEM_NAME` as `csname`,`ntag`.`tag`,`ntag`.`descr`,alsm_license_info .DESCRIPTION,`mqt`.`ntag`,
		`alsm_license_info`.`PN`,`alsm_license_info`.`DESCRIPTION` as pdescription, `info`.`ACTIVATION_ON` `aon`, `conv`.`mqt_product`, `conv`.`mqt_config`, `mqt`.`DISC`, `mqt`.`TS_WAR`,
		`ROOT_LICENSE_ID_ACTIVATION_DATE`,`conv`.`mqt_u_name`
		FROM `alsm_gen_info` as `info`
		right JOIN alsm_license_info ON `info`.LICENSE_ID  = alsm_license_info.LICENSE_ID
		left join `alsm_cust_profile` as `mqt` on `mqt`.`ROOT_LICENSE_ID`=`info`.`ROOT_LICENSE_ID` 
		left join `ntag` as `ntag` on `ntag`.`id`=`mqt`.`ntag` 
		left join `shelf_mqt_conv` as `conv` on `conv`.`alsm_name`=`info`.`PRODUCT_NAME`		
		Where `mqt`.`ntag`='".cib_dbEscape($ticket166['tag'])."' GROUP BY `info`.`PRODUCT_NAME`,`info`.`PRODUCT_RELEASE_NAME`,`alsm_license_info`.`PN`";
	
		$res55 = cib_dbQuery($sql55);
		while ($ticket55=cib_dbFetchAssoc($res55))
		{
		
		$sql56_3 = "select COUNT(DISTINCT `info`.`ROOT_LICENSE_ID`)
		FROM `alsm_gen_info` as `info`
		right JOIN alsm_license_info ON `info`.LICENSE_ID  = alsm_license_info.LICENSE_ID
		left join `alsm_cust_profile` as `mqt` on `mqt`.`ROOT_LICENSE_ID`=`info`.`ROOT_LICENSE_ID` 
		left join `ntag` as `ntag` on `ntag`.`id`=`mqt`.`ntag` 
		Where `mqt`.`ntag`='".cib_dbEscape($ticket55['ntag'])."' 
		AND `info`.`PRODUCT_RELEASE_NAME`='".cib_dbEscape($ticket55['release'])."' AND `info`.`PRODUCT_NAME`='".cib_dbEscape($ticket55['product'])."' ";		
		//echo $sql56_3.'<br>';
		$res56_3 = cib_dbQuery($sql56_3);
		$ticket55['rlid']=cib_dbResult($res56_3);	
		
		
		$sql56 = "SELECT SUM(`alsm_license_info`.QTY_OF_PN) as sum_parts
		FROM `alsm_gen_info` as `info`
		right JOIN alsm_license_info ON `info`.LICENSE_ID  = alsm_license_info.LICENSE_ID
		left join `alsm_cust_profile` as `mqt` on `mqt`.`ROOT_LICENSE_ID`=`info`.`ROOT_LICENSE_ID` 
		left join `ntag` as `ntag` on `ntag`.`id`=`mqt`.`ntag` 
		Where `mqt`.`ntag`='".cib_dbEscape($ticket55['ntag'])."' AND `alsm_license_info`.`PN`='".cib_dbEscape($ticket55['PN'])."' 
		AND `info`.`PRODUCT_RELEASE_NAME`='".cib_dbEscape($ticket55['release'])."' AND `info`.`PRODUCT_NAME`='".cib_dbEscape($ticket55['product'])."' ";
	
		$res56 = cib_dbQuery($sql56);
		$ticket55['sum_parts']=cib_dbResult($res56);
		
		
		$objPHPExcel->getActiveSheet()->setCellValue('B'.$i, $cnt)->getStyle('B'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('C'.$i, '')->getStyle('C'.$i)->applyFromArray($styleArray01);		
		$objPHPExcel->getActiveSheet()->setCellValue('D'.$i, $ntag_n_d)->getStyle('D'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('E'.$i, $ticket55['PN'])->getStyle('E'.$i)->applyFromArray($styleArray01);	
		$objPHPExcel->getActiveSheet()->setCellValue('F'.$i, $ticket55['pdescription'])->getStyle('F'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('G'.$i, $ticket55['product'])->getStyle('G'.$i)->applyFromArray($styleArray01);		
		$objPHPExcel->getActiveSheet()->setCellValue('H'.$i, $ticket55['sum_parts'])->getStyle('H'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('I'.$i, '')->getStyle('I'.$i)->applyFromArray($styleArray01);	
		$objPHPExcel->getActiveSheet()->setCellValue('J'.$i, '')->getStyle('J'.$i)->applyFromArray($styleArray01);			
		
		$bdescr='';	
		$unitprice=0;
		$tunitprice=0;
		$mqtval=0;
		$mqtvaltotal=0;
		$mqtname='MQT Unit not available';
		$cur='-';
		$sq22 = "SELECT * FROM `tsunits` WHERE 	item='".cib_dbEscape($ticket55['PN'])."' LIMIT 1";	
		//echo $sq22.'<BR>';
		$res22 = cib_dbQuery($sq22);
		$bordch22=cib_dbFetchAssoc($res22);
		$sq23 = "SELECT * FROM `tsunits` WHERE 	item='".substr_replace(cib_dbEscape($ticket55['PN']),'%',8,-1)."' LIMIT 1";	
		//echo $sq23.'<BR>';
		$res23 = cib_dbQuery($sq23);
		$bordch23=cib_dbFetchAssoc($res23);
		$sq24 = "SELECT * FROM `tsunits` WHERE 	item='".substr_replace(cib_dbEscape($ticket55['PN']),'%%',8,2)."' LIMIT 1";	
		//echo $sq24.'<BR>';
		$res24 = cib_dbQuery($sq24);
		$bordch24=cib_dbFetchAssoc($res24);
		//echo $sq24.'<br>';
		if ($bordch22){
		//echo $sq22.'<br>';
		$bdescr=$bordch22['descr'];
		$unitprice=$bordch22['price'];
		$mqtval=$bordch22['valu'];
		$mqtname=$bordch22['unit'];
		$cur=$bordch22['cur'];
		}						
		else if ($bordch23 AND !$bordch22){
		//echo $sq23.'<br>';
		$bdescr=$bordch23['descr'];
		$unitprice=$bordch23['price'];
		$mqtval=$bordch23['valu'];
		$mqtname=$bordch23['unit'];
		$cur=$bordch23['cur'];
		}	
		else if ($bordch24 AND !$bordch22 AND !$bordch23){
		//echo $sq23.'<br>';
		$bdescr=$bordch24['descr'];
		$unitprice=$bordch24['price'];
		$mqtval=$bordch24['valu'];
		$mqtname=$bordch24['unit'];
		$cur=$bordch24['cur'];
		}
		
		$mqtname=$ticket55['mqt_u_name'];
		
		if ($cur=='USD' AND $runc['curr']=='EUR') $unitprice=$unitprice*$USD_EUR;
		else if ($cur=='EUR' AND $runc['curr']=='USD') $unitprice=$unitprice*$EUR_USD;
		
		$tunitprice=$ticket55['sum_parts']*$unitprice;
		$mqtvaltotal=$ticket55['sum_parts']*$mqtval;
		
		$objPHPExcel->getActiveSheet()->setCellValue('K'.$i, $ticket55['mqt_product'])->getStyle('K'.$i)->applyFromArray($styleArray01);	
		$objPHPExcel->getActiveSheet()->setCellValue('L'.$i, $ticket55['mqt_config'])->getStyle('L'.$i)->applyFromArray($styleArray01);		
		$objPHPExcel->getActiveSheet()->setCellValue('N'.$i, $unitprice)->getStyle('N'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('O'.$i, $tunitprice)->getStyle('O'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->getStyle('P'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);	
		$objPHPExcel->getActiveSheet()->setCellValue('P'.$i, '='.$ticket55['DISC']/100)->getStyle('P'.$i)->applyFromArray($styleArray01);					
		$objPHPExcel->getActiveSheet()->setCellValue('Q'.$i, '')->getStyle('Q'.$i)->applyFromArray($styleArray01);				
		$objPHPExcel->getActiveSheet()->setCellValue('R'.$i, '=N'.$i.'-(N'.$i.'*P'.$i.')')->getStyle('R'.$i)->applyFromArray($styleArray01);					
		$objPHPExcel->getActiveSheet()->setCellValue('S'.$i, '=R'.$i.'*H'.$i.'')->getStyle('S'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('M'.$i, $ticket55['release'])->getStyle('M'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('T'.$i, $mqtval)->getStyle('T'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('U'.$i, $mqtvaltotal)->getStyle('U'.$i)->applyFromArray($styleArray01);		
		$objPHPExcel->getActiveSheet()->setCellValue('V'.$i, $mqtname)->getStyle('V'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('W'.$i, ' '.yearfm($ticket55['TS_WAR']))->getStyle('W'.$i)->applyFromArray($styleArray01);
		
		$alsm_ts_war=$ticket55['TS_WAR'];
		
		
		$sql56_1 = "SELECT IFNULL(SUM(`alsm_license_info`.QTY_OF_PN),0)
		FROM `alsm_gen_info` as `info`
		right JOIN alsm_license_info ON `info`.LICENSE_ID  = alsm_license_info.LICENSE_ID
		left join `alsm_cust_profile` as `mqt` on `mqt`.`ROOT_LICENSE_ID`=`info`.`ROOT_LICENSE_ID` 
		left join `ntag` as `ntag` on `ntag`.`id`=`mqt`.`ntag` 
		Where `mqt`.`ntag`='".cib_dbEscape($ticket55['ntag'])."' AND `alsm_license_info`.`PN`='".cib_dbEscape($ticket55['PN'])."' 
		AND `info`.`PRODUCT_RELEASE_NAME`='".cib_dbEscape($ticket55['release'])."' AND `info`.`PRODUCT_NAME`='".cib_dbEscape($ticket55['product'])."' 
		AND ADDDATE(`info`.`ROOT_LICENSE_ID_ACTIVATION_DATE`, INTERVAL $alsm_ts_war MONTH) > '$mtcpst_o' AND `info`.`ROOT_LICENSE_ID_ACTIVATION_DATE` < '$mtcpst_o' ";
		//echo $sql56_1."<br>";
		$res56_1 = cib_dbQuery($sql56_1);
		$Alsm_in_war_prio = 0;
		$Alsm_in_war_prio = cib_dbResult($res56_1);
		
		$sql56_2 = "SELECT IFNULL(SUM(`alsm_license_info`.QTY_OF_PN),0)
		FROM `alsm_gen_info` as `info`
		right JOIN alsm_license_info ON `info`.LICENSE_ID  = alsm_license_info.LICENSE_ID
		left join `alsm_cust_profile` as `mqt` on `mqt`.`ROOT_LICENSE_ID`=`info`.`ROOT_LICENSE_ID` 
		left join `ntag` as `ntag` on `ntag`.`id`=`mqt`.`ntag` 
		Where `mqt`.`ntag`='".cib_dbEscape($ticket55['ntag'])."' AND `alsm_license_info`.`PN`='".cib_dbEscape($ticket55['PN'])."' 
		AND `info`.`PRODUCT_RELEASE_NAME`='".cib_dbEscape($ticket55['release'])."' AND `info`.`PRODUCT_NAME`='".cib_dbEscape($ticket55['product'])."' 
		AND ADDDATE(`info`.`ROOT_LICENSE_ID_ACTIVATION_DATE`, INTERVAL $alsm_ts_war MONTH) <= '$mtcpst_o' AND `info`.`ROOT_LICENSE_ID_ACTIVATION_DATE` < '$mtcpst_o' ";
		//echo $sql56_2."<br>";
		$res56_2 = cib_dbQuery($sql56_2);			
		$Alsm_oo_war_prio = 0;
		$Alsm_oo_war_prio = cib_dbResult($res56_2);		
		
		
		$objPHPExcel->getActiveSheet()->setCellValue('X'.$i, $Alsm_in_war_prio)->getStyle('X'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('Y'.$i, $Alsm_oo_war_prio)->getStyle('Y'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('Z'.$i, $Alsm_in_war_prio+$Alsm_oo_war_prio)->getStyle('Z'.$i)->applyFromArray($styleArray01);
		
		$objPHPExcel->getActiveSheet()->setCellValue('AA'.$i, '=X'.$i.'*R'.$i.'')->getStyle('AA'.$i)->applyFromArray($styleArray01);					
		$objPHPExcel->getActiveSheet()->setCellValue('AB'.$i, '=Y'.$i.'*R'.$i.'')->getStyle('AB'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('AC'.$i, '=AA'.$i.'+AB'.$i.'')->getStyle('AC'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->getStyle('AC'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);
		
		$objPHPExcel->getActiveSheet()->setCellValue('AD'.$i, '=X'.$i.'*T'.$i.'')->getStyle('AD'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('AE'.$i, '=Y'.$i.'*T'.$i.'')->getStyle('AE'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('AF'.$i, '=AD'.$i.'+AE'.$i.'')->getStyle('AF'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('AG'.$i, $mqtname)->getStyle('AG'.$i)->applyFromArray($styleArray01);		
		
		
		$sql56_4 = "SELECT IFNULL(SUM(`alsm_license_info`.QTY_OF_PN),0)
		FROM `alsm_gen_info` as `info`
		right JOIN alsm_license_info ON `info`.LICENSE_ID  = alsm_license_info.LICENSE_ID
		left join `alsm_cust_profile` as `mqt` on `mqt`.`ROOT_LICENSE_ID`=`info`.`ROOT_LICENSE_ID` 
		left join `ntag` as `ntag` on `ntag`.`id`=`mqt`.`ntag` 
		Where `mqt`.`ntag`='".cib_dbEscape($ticket55['ntag'])."' AND `alsm_license_info`.`PN`='".cib_dbEscape($ticket55['PN'])."' 
		AND `info`.`PRODUCT_RELEASE_NAME`='".cib_dbEscape($ticket55['release'])."' AND `info`.`PRODUCT_NAME`='".cib_dbEscape($ticket55['product'])."' 
		AND ADDDATE(`info`.`ROOT_LICENSE_ID_ACTIVATION_DATE`, INTERVAL $alsm_ts_war MONTH) > '$mtcstd_o' AND `info`.`ROOT_LICENSE_ID_ACTIVATION_DATE` < '$mtcstd_o' ";
		//echo $sql56_1."<br>";
		$res56_4 = cib_dbQuery($sql56_4);
		$Alsm_in_war_prio_start = 0;
		$Alsm_in_war_prio_start = cib_dbResult($res56_4);		
		
		
		
		$sql56_3 = "SELECT IFNULL(SUM(`alsm_license_info`.QTY_OF_PN),0)
		FROM `alsm_gen_info` as `info`
		right JOIN alsm_license_info ON `info`.LICENSE_ID  = alsm_license_info.LICENSE_ID
		left join `alsm_cust_profile` as `mqt` on `mqt`.`ROOT_LICENSE_ID`=`info`.`ROOT_LICENSE_ID` 
		left join `ntag` as `ntag` on `ntag`.`id`=`mqt`.`ntag` 
		Where `mqt`.`ntag`='".cib_dbEscape($ticket55['ntag'])."' AND `alsm_license_info`.`PN`='".cib_dbEscape($ticket55['PN'])."' 
		AND `info`.`PRODUCT_RELEASE_NAME`='".cib_dbEscape($ticket55['release'])."' AND `info`.`PRODUCT_NAME`='".cib_dbEscape($ticket55['product'])."' 
		AND ADDDATE(`info`.`ROOT_LICENSE_ID_ACTIVATION_DATE`, INTERVAL $alsm_ts_war MONTH) <= '$mtcstd_o' AND `info`.`ROOT_LICENSE_ID_ACTIVATION_DATE` < '$mtcstd_o' ";
		//echo $sql56_1."<br>";
		$res56_3 = cib_dbQuery($sql56_3);			
		$Alsm_oo_war_prio_start = 0;
		$Alsm_oo_war_prio_start = cib_dbResult($res56_3);			
		
		$objPHPExcel->getActiveSheet()->setCellValue('AH'.$i, $Alsm_in_war_prio_start)->getStyle('AH'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('AI'.$i, $Alsm_oo_war_prio_start)->getStyle('AI'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('AJ'.$i, $Alsm_in_war_prio_start+$Alsm_oo_war_prio_start)->getStyle('AJ'.$i)->applyFromArray($styleArray01);
		
		$objPHPExcel->getActiveSheet()->setCellValue('AK'.$i, '=AH'.$i.'*R'.$i.'')->getStyle('AK'.$i)->applyFromArray($styleArray01);					
		$objPHPExcel->getActiveSheet()->setCellValue('AL'.$i, '=AI'.$i.'*R'.$i.'')->getStyle('AL'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('AM'.$i, '=AK'.$i.'+AL'.$i.'')->getStyle('AM'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->getStyle('AM'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);
		
		$objPHPExcel->getActiveSheet()->setCellValue('AN'.$i, '=AH'.$i.'*T'.$i.'')->getStyle('AN'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('AO'.$i, '=AI'.$i.'*T'.$i.'')->getStyle('AO'.$i)->applyFromArray($styleArray01);
		$objPHPExcel->getActiveSheet()->setCellValue('AP'.$i, '=AN'.$i.'+AO'.$i.'')->getStyle('AP'.$i)->applyFromArray($styleArray01);
		
		$alsm_prod=$ticket55['product'];
		$alsm_prod_rel=$ticket55['release'];
			
		$Node_collector_alsm[]=$ntag_n_d.':'.$alsm_prod.':'.$alsm_prod_rel;
		//$Node_collector_alsm_nodes[$ntag_n_d.':'.$alsm_prod.':'.$alsm_prod_rel][]=$ticket55['rlid'];
		$Node_collector_alsm_nodes[$ntag_n_d.':'.$alsm_prod.':'.$alsm_prod_rel] =$ticket55['rlid'];		
		$Node_collector_alsm_mqt_prod[$ntag_n_d.':'.$alsm_prod.':'.$alsm_prod_rel][]=$ticket55['mqt_product'];
		$Node_collector_alsm_mqt_config[$ntag_n_d.':'.$alsm_prod.':'.$alsm_prod_rel][]=$ticket55['mqt_config'];
		$Node_collector_alsm_q_parts[$ntag_n_d.':'.$alsm_prod.':'.$alsm_prod_rel][]=$ticket55['sum_parts'];
		$Node_collector_alsm_mqtvaltotal[$ntag_n_d.':'.$alsm_prod.':'.$alsm_prod_rel][]=$mqtvaltotal;
		$Node_collector_alsm_mqtname[$ntag_n_d.':'.$alsm_prod.':'.$alsm_prod_rel][]=$mqtname;
		$Node_collector_alsm_glp_price_total[$ntag_n_d.':'.$alsm_prod.':'.$alsm_prod_rel][]=$tunitprice;
		$Node_collector_alsm_psp_price_total[$ntag_n_d.':'.$alsm_prod.':'.$alsm_prod_rel][]=($tunitprice-($tunitprice*($ticket55['DISC']/100)));
		$Node_collector_alsm_disc[$ntag_n_d.':'.$alsm_prod.':'.$alsm_prod_rel][]=$ticket55['DISC']/100;
		
		$mqt_product=$ticket55['mqt_product'];
		$mqt_config=$ticket55['mqt_config'];
		$reldisp=$alsm_prod_rel;
		
		$MQT_ALSM_SUM[]=$mqt_product.':'.$mqt_config.':'.$reldisp;
		$MQT_ALSM_SUM_mqtname[$mqt_product.':'.$mqt_config.':'.$reldisp]=$mqtname;
		$MQT_ALSM_SUM_MQT_Unit_Qty_Prio[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($Alsm_in_war_prio+$Alsm_oo_war_prio)*$mqtval;
		$MQT_ALSM_AA_FP[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($Alsm_in_war_prio+$Alsm_oo_war_prio)*$unitprice;
		$MQT_ALSM_AA[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($Alsm_in_war_prio+$Alsm_oo_war_prio)*($unitprice-($unitprice*($ticket55['DISC']/100)));
		$MQT_ALSM_AB[$mqt_product.':'.$mqt_config.':'.$reldisp][]=$Alsm_in_war_prio*$mqtval;
		$MQT_ALSM_AC[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($Alsm_oo_war_prio)*$mqtval;
		$MQT_ALSM_Z_FP[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($Alsm_oo_war_prio)*$unitprice;
		$MQT_ALSM_BQ[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($Alsm_oo_war_prio)*$mqtval;
		$MQT_ALSM_AK_FP[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($Alsm_in_war_prio_start+$Alsm_oo_war_prio_start)*$unitprice;
		$MQT_ALSM_AK[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($Alsm_in_war_prio_start+$Alsm_oo_war_prio_start)*($unitprice-($unitprice*($ticket55['DISC']/100)));
		$MQT_ALSM_AN[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($Alsm_in_war_prio_start+$Alsm_oo_war_prio_start)*$mqtval;
		$MQT_ALSM_AL[$mqt_product.':'.$mqt_config.':'.$reldisp][]=$Alsm_in_war_prio_start*$mqtval;
		$MQT_ALSM_AN2[$mqt_product.':'.$mqt_config.':'.$reldisp][]=$Alsm_oo_war_prio_start*$mqtval;
		$MQT_ALSM_AO[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($Alsm_oo_war_prio_start)*($unitprice-($unitprice*($ticket55['DISC']/100)));
		
		$i++;
		$cnt++;	
		}
		
	}
$cnt=$cnt+7;
$objPHPExcel->getActiveSheet()->setAutoFilter('B7:W7');
If ($runc['curr']=='EUR') 
{

	$objPHPExcel->getActiveSheet()->getStyle('O'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('S'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AB'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AC'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AA'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AK'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AL'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AM'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
				
	}
else 
{	
	$objPHPExcel->getActiveSheet()->getStyle('O'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('S'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AB'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AC'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AA'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AK'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AL'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AM'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);

}
$objPHPExcel->getActiveSheet()->setCellValue('O'.$cnt, '=SUBTOTAL(9,O8:O'.($cnt-1).')')->getStyle('O'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('S'.$cnt, '=SUBTOTAL(9,S8:S'.($cnt-1).')')->getStyle('S'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('X'.$cnt, '=SUBTOTAL(9,X8:X'.($cnt-1).')')->getStyle('X'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('Y'.$cnt, '=SUBTOTAL(9,Y8:Y'.($cnt-1).')')->getStyle('Y'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('Z'.$cnt, '=SUBTOTAL(9,Z8:Z'.($cnt-1).')')->getStyle('Z'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AA'.$cnt, '=SUBTOTAL(9,AA8:AA'.($cnt-1).')')->getStyle('AA'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AB'.$cnt, '=SUBTOTAL(9,AB8:AB'.($cnt-1).')')->getStyle('AB'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AC'.$cnt, '=SUBTOTAL(9,AC8:AC'.($cnt-1).')')->getStyle('AC'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AD'.$cnt, '=SUBTOTAL(9,AD8:AD'.($cnt-1).')')->getStyle('AD'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AE'.$cnt, '=SUBTOTAL(9,AE8:AE'.($cnt-1).')')->getStyle('AE'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AF'.$cnt, '=SUBTOTAL(9,AF8:AF'.($cnt-1).')')->getStyle('AF'.$cnt)->applyFromArray($styleArray01);

$objPHPExcel->getActiveSheet()->setCellValue('AH'.$cnt, '=SUBTOTAL(9,AH8:AH'.($cnt-1).')')->getStyle('AH'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AI'.$cnt, '=SUBTOTAL(9,AI8:AI'.($cnt-1).')')->getStyle('AI'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AJ'.$cnt, '=SUBTOTAL(9,AJ8:AJ'.($cnt-1).')')->getStyle('AJ'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AK'.$cnt, '=SUBTOTAL(9,AK8:AK'.($cnt-1).')')->getStyle('AK'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AL'.$cnt, '=SUBTOTAL(9,AL8:AL'.($cnt-1).')')->getStyle('AL'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AM'.$cnt, '=SUBTOTAL(9,AM8:AM'.($cnt-1).')')->getStyle('AM'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AN'.$cnt, '=SUBTOTAL(9,AN8:AN'.($cnt-1).')')->getStyle('AN'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AO'.$cnt, '=SUBTOTAL(9,AO8:AO'.($cnt-1).')')->getStyle('AO'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AP'.$cnt, '=SUBTOTAL(9,AP8:AP'.($cnt-1).')')->getStyle('AP'.$cnt)->applyFromArray($styleArray01);
}

//print_r(array_unique($Node_collector_alsm));
$i=8;
$cnt=1;

If ($is_el_tr) {
$tab_cnt++;
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex($tab_cnt)->setTitle("Electronic Audit Details")->getTabColor()->setRGB('FF0000');
$tabsnames[]="Electronic Audit Details";
$objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(85);
$objPHPExcel->getActiveSheet()->setCellValue('A1', "Care Renewal Assessment (CRA)")->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('A2', "Electronic Audit Details")->getStyle('A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('A2')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->getStyle('A2')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('A3', "This table lists the Part by part breakdown found for any products that Sonar's Audit Source is from an NMS audit  (see Column A in Summary tab)")->getStyle('A3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('A3')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('J1', "**** Changes made to this tab will NOT be reflected in the Summary tab ****")->getStyle('J1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('J1')->getFont()->setSize(12)->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->getStyle('J1')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->mergeCells('J1:L1');

$objPHPExcel->getActiveSheet()->setCellValue('V3', "Technical Support (TS)")->getStyle('V3');
$objPHPExcel->getActiveSheet()->setCellValue('Z2', "Prior Care Contract start date: $mtcpst")->getStyle('Z2');
$objPHPExcel->getActiveSheet()->setCellValue('Z3', "New Care Contract start date: $mtcstd")->getStyle('Z2');
$objPHPExcel->getActiveSheet()->setCellValue('BJ3', "RES/HWS")->getStyle('BJ3');
$objPHPExcel->getActiveSheet()->setCellValue('BN2', "Prior Care Contract start date: $mtcpst")->getStyle('BN2');
$objPHPExcel->getActiveSheet()->setCellValue('BN3', "New Care Contract start date: $mtcstd")->getStyle('BN2');
$objPHPExcel->getActiveSheet()->getColumnDimension('V')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('w')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('X')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AC')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AD')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AE')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AF')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AG')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AH')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AI')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AJ')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AK')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AL')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AM')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AN')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AO')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AP')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AQ')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AR')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AS')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AT')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AU')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AN')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AV')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AW')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AX')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AY')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AZ')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BA')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BB')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BC')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BD')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BE')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BF')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BG')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BH')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BJ')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BK')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BL')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BM')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BN')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BO')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BP')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BQ')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BR')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BS')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BT')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BU')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BV')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BW')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BX')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BY')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BZ')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CA')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CB')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CC')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CD')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CE')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CF')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CG')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AT')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CH')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CI')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CJ')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CK')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CL')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CM')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CN')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CO')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CP')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CQ')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CR')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CS')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CT')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CU')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CV')->setOutlineLevel(1)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AO')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AP')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AQ')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AR')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AS')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AT')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AU')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AV')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AW')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AX')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AY')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AZ')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BA')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BB')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BC')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BD')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BE')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BF')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BG')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CC')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CD')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CE')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CF')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CG')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CH')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CI')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CJ')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CK')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CL')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CM')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CN')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CO')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CP')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CQ')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CR')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CS')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CT')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('CU')->setOutlineLevel(2)->setVisible(true)->setCollapsed(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('B7', "Item")->getStyle('B7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('C6', "Note: To see details about 1830 PSS shelf type and quantity, please filter 'IBR' in the 'Part Number' column. 'Part Description' column will describe shelf type and 'Qty of parts' count of the shelves.")->getStyle('C6')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('C7', 'Part Number')->getStyle('C7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('D7', 'Part Description')->getStyle('D7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('E7', 'Shelf type (name in NMS)')->getStyle('E7')->applyFromArray($styleArray0);	
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('F7', 'Qty of parts')->getStyle('F7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('G7', 'NTAG')->getStyle('G7')->applyFromArray($styleArray0);	
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('H7', 'Warning')->getStyle('H7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('I7', 'MQT Product Name')->getStyle('I7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('J7', 'MQT Product Config')->getStyle('J7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('K7', 'SW Release(s)')->getStyle('K7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('L6', "1 USD = ".$USD_EUR." EUR")->getStyle('L6')->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('L7', "List Price (".$runc['curr'].") \nEACH")->getStyle('L7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('M7', "List Price (".$runc['curr'].")\nTOTAL")->getStyle('M7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('N7', "Customer\nDiscount")->getStyle('N7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('O7', "Confidence in\ncustomer\nDiscount")->getStyle('O7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('P')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('P7', 'PSP ('.$runc['curr'].') Each ')->getStyle('P7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('Q7', "PSP (".$runc['curr'].") Total")->getStyle('Q7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('R7', "MQT Unit\nEACH")->getStyle('R7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('S')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('S7', "MQT Unit\nTOTAL")->getStyle('S7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('T')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('T7', "MQT Unit Name")->getStyle('T7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('U')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('U7', "TS Warranty\nLength")->getStyle('U7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->mergeCells('V5:AE5');
$objPHPExcel->getActiveSheet()->setCellValue('V5', "Historical View. Deployed as of Prior Contract Date as of $mtcpst")->getStyle('V5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('W5', "")->getStyle('W5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('X5', "")->getStyle('X5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('Y5', "")->getStyle('Y5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('Z5', "")->getStyle('Z5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AA5', "")->getStyle('AA5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AB5', "")->getStyle('AB5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AC5', "")->getStyle('AC5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AD5', "")->getStyle('AD5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AE5', "")->getStyle('AE5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->mergeCells('V6:X6');
$objPHPExcel->getActiveSheet()->setCellValue('V6', "Part Number Quantities")->getStyle('V6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('W6', "")->getStyle('W6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('X6', "")->getStyle('X6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('V')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('V7', "IW")->getStyle('V7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('W7', "OOW")->getStyle('W7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('X7', "Total")->getStyle('X7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->mergeCells('Y6:AA6');
$objPHPExcel->getActiveSheet()->setCellValue('Y6', "Product Selling Price (PSP) values (".$runc['curr'].")")->getStyle('Y6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('Z6', "")->getStyle('Z6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AA6', "")->getStyle('AA6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('Y7', "IW")->getStyle('Y7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('Z7', "OOW")->getStyle('Z7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AA7', "Total")->getStyle('AA7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->mergeCells('AB6:AE6');
$objPHPExcel->getActiveSheet()->setCellValue('AB6', "MQT Unit Quantities")->getStyle('AB6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AC6', "")->getStyle('AC6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AD6', "")->getStyle('AD6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AE6', "")->getStyle('AE6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AB7', "IW")->getStyle('AB7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('AC')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AC7', "OOW")->getStyle('AC7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('AD')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AD7', "Total")->getStyle('AD7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('AE')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('AE7', "MQT Unit Name")->getStyle('AE7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->mergeCells('AF6:AH6');
$objPHPExcel->getActiveSheet()->setCellValue('AF6', "Part Number Quantities")->getStyle('AF6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AG6', "")->getStyle('AG6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AH6', "")->getStyle('AH6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AF')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AF7', "IW")->getStyle('AF7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AG')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AG7', "OOW")->getStyle('AG7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AH')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AH7', "Total")->getStyle('AH7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('AI6:AK6');
$objPHPExcel->getActiveSheet()->setCellValue('AI6', "Product Selling Price (PSP) values (".$runc['curr'].")")->getStyle('AI6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AJ6', "")->getStyle('AJ6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AK6', "")->getStyle('AK6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AI')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AI7', "IW")->getStyle('AI7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AJ')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AJ7', "OOW")->getStyle('AJ7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AK')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AK7', "Total")->getStyle('AK7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('AL6:AN6');
$objPHPExcel->getActiveSheet()->setCellValue('AL6', "MQT Unit Quantities")->getStyle('AL6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AM6', "")->getStyle('AM6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AN6', "")->getStyle('AN6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AL')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AL7', "IW")->getStyle('AL7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AM')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AM7', "OOW")->getStyle('AM7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AN')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AN7', "Total")->getStyle('AN7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('AF5:AN5');
$objPHPExcel->getActiveSheet()->setCellValue('AF5', "New Contract View: 1st year as of $mtcstd")->getStyle('AF5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AG5', "")->getStyle('AG5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AH5', "")->getStyle('AH5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AI5', "")->getStyle('AI5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AJ5', "")->getStyle('AJ5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AK5', "")->getStyle('AK5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AL5', "")->getStyle('AL5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AM5', "")->getStyle('AM5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AN5', "")->getStyle('AN5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AO')->setWidth('1');
$objPHPExcel->getActiveSheet()->setCellValue('AO5', "")->getStyle('AO5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AO6', "")->getStyle('AO6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AO7', "")->getStyle('AO7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('AP6:AR6');
$objPHPExcel->getActiveSheet()->setCellValue('AP6', "Part Number Quantities")->getStyle('AP6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AQ6', "")->getStyle('AQ6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AR6', "")->getStyle('AR6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AP')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AP7', "IW")->getStyle('AP7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AQ')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AQ7', "OOW")->getStyle('AQ7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AR')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AR7', "Total")->getStyle('AR7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('AS6:AU6');
$objPHPExcel->getActiveSheet()->setCellValue('AS6', "Product Selling Price (PSP) values (".$runc['curr'].")")->getStyle('AS6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AT6', "")->getStyle('AT6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AU6', "")->getStyle('AU6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AS')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AS7', "IW")->getStyle('AS7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AT')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AT7', "OOW")->getStyle('AT7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AU')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AU7', "Total")->getStyle('AU7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('AV6:AX6');
$objPHPExcel->getActiveSheet()->setCellValue('AV6', "MQT Unit Quantities")->getStyle('AV6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AW6', "")->getStyle('AW6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AX6', "")->getStyle('AX6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AV')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AV7', "IW")->getStyle('AV7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AW')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AW7', "OOW")->getStyle('AW7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AX')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AX7', "Total")->getStyle('AX7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('AP5:AX5');
$objPHPExcel->getActiveSheet()->setCellValue('AP5', "New Contract View: 2nd year as of $mtcstd_1")->getStyle('AP5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AQ5', "")->getStyle('AQ5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AR5', "")->getStyle('AR5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AS5', "")->getStyle('AS5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AT5', "")->getStyle('AT5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AU5', "")->getStyle('AU5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AV5', "")->getStyle('AV5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AW5', "")->getStyle('AW5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AX5', "")->getStyle('AX5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('AY6:BA6');
$objPHPExcel->getActiveSheet()->setCellValue('AY6', "Part Number Quantities")->getStyle('AY6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AZ6', "")->getStyle('AZ6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BA6', "")->getStyle('BA6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AY')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AY7', "IW")->getStyle('AY7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AZ')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('AZ7', "OOW")->getStyle('AZ7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BA')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('BA7', "Total")->getStyle('BA7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('BB6:BD6');
$objPHPExcel->getActiveSheet()->setCellValue('BB6', "Product Selling Price (PSP) values (".$runc['curr'].")")->getStyle('BB6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BC6', "")->getStyle('BC6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BD6', "")->getStyle('BD6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BB')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('BB7', "IW")->getStyle('BB7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BC')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('BC7', "OOW")->getStyle('BC7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BD')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('BD7', "Total")->getStyle('BD7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('BE6:BG6');
$objPHPExcel->getActiveSheet()->setCellValue('BE6', "MQT Unit Quantities")->getStyle('BE6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BF6', "")->getStyle('BF6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BG6', "")->getStyle('BG6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BE')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('BE7', "IW")->getStyle('BE7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BF')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('BF7', "OOW")->getStyle('BF7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BG')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('BG7', "Total")->getStyle('BG7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('AY5:BG5');
$objPHPExcel->getActiveSheet()->setCellValue('AY5', "New Contract View: 3rd year as of $mtcstd_2")->getStyle('AY5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AZ5', "")->getStyle('AZ5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BA5', "")->getStyle('BA5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BB5', "")->getStyle('BB5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BC5', "")->getStyle('BC5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BD5', "")->getStyle('BD5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BE5', "")->getStyle('BE5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BF5', "")->getStyle('BF5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BG5', "")->getStyle('BG5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BH')->setWidth('2');
$objPHPExcel->getActiveSheet()->setCellValue('BH5', "")->getStyle('BH5')->applyFromArray($styleArray03);
$objPHPExcel->getActiveSheet()->setCellValue('BH6', "")->getStyle('BH6')->applyFromArray($styleArray03);
$objPHPExcel->getActiveSheet()->setCellValue('BH7', "")->getStyle('BH7')->applyFromArray($styleArray03);
$objPHPExcel->getActiveSheet()->setCellValue('BI7', "RES/HWS\nWarranty")->getStyle('BI7')->applyFromArray($styleArray1);
$objPHPExcel->getActiveSheet()->getColumnDimension('BI')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->mergeCells('BJ6:BL6');
$objPHPExcel->getActiveSheet()->setCellValue('BJ6', "Part Number Quantities")->getStyle('BJ6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BK6', "")->getStyle('BK6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BL6', "")->getStyle('BL6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('BJ')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('BJ7', "IW")->getStyle('BJ7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('BK')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('BK7', "OOW")->getStyle('BK7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('BL')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('BL7', "Total")->getStyle('BL7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->mergeCells('BM6:BO6');
$objPHPExcel->getActiveSheet()->setCellValue('BM6', "Product Selling Price (PSP) values (".$runc['curr'].")")->getStyle('BM6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BN6', "")->getStyle('BN6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BO6', "")->getStyle('BO6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->mergeCells('BM6:BO6');
$objPHPExcel->getActiveSheet()->setCellValue('BM6', "Product Selling Price (PSP) values (".$runc['curr'].")")->getStyle('BM6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BN6', "")->getStyle('BN6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BO6', "")->getStyle('BO6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('BM')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('BM7', "IW")->getStyle('BM7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('BN')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('BN7', "OOW")->getStyle('BN7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('BO')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('BO7', "Total")->getStyle('BO7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->mergeCells('BP6:BS6');
$objPHPExcel->getActiveSheet()->setCellValue('BP6', "MQT Unit Quantities")->getStyle('BP6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BQ6', "")->getStyle('BQ6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BR6', "")->getStyle('BR6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BS6', "")->getStyle('BS6')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('BP')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('BP7', "IW")->getStyle('BP7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('BQ')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('BQ7', "OOW")->getStyle('BQ7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('BR')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('BR7', "Total")->getStyle('BR7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->getColumnDimension('BS')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('BS7', "MQT Unit Name")->getStyle('BS7')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->mergeCells('BJ5:BS5');
$objPHPExcel->getActiveSheet()->setCellValue('BJ5', "Historical View. Deployed as of Prior Contract Date as of $mtcpst")->getStyle('BJ5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BK5', "")->getStyle('BK5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BL5', "")->getStyle('BL5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BM5', "")->getStyle('BM5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BN5', "")->getStyle('BN5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BO5', "")->getStyle('BO5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BP5', "")->getStyle('BP5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BQ5', "")->getStyle('BQ5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BR5', "")->getStyle('BR5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('BS5', "")->getStyle('BS5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->setCellValue('AE5', "")->getStyle('AE5')->applyFromArray($styleArray00);
$objPHPExcel->getActiveSheet()->mergeCells('BT6:BV6');
$objPHPExcel->getActiveSheet()->setCellValue('BT6', "Part Number Quantities")->getStyle('BT6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BU6', "")->getStyle('BU6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BV6', "")->getStyle('BV6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BT')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('BT7', "IW")->getStyle('BT7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BU')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('BU7', "OOW")->getStyle('BU7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BV')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('BV7', "Total")->getStyle('BV7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('BW6:BY6');
$objPHPExcel->getActiveSheet()->setCellValue('BW6', "Product Selling Price (PSP) values (".$runc['curr'].")")->getStyle('BW6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BX6', "")->getStyle('BX6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BY6', "")->getStyle('BY6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BW')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('BW7', "IW")->getStyle('BW7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BX')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('BX7', "OOW")->getStyle('BX7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BY')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('BY7', "Total")->getStyle('BY7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('BZ6:CB6');
$objPHPExcel->getActiveSheet()->setCellValue('BZ6', "MQT Unit Quantities")->getStyle('BZ6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CA6', "")->getStyle('CA6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CB6', "")->getStyle('CB6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BZ')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('BZ7', "IW")->getStyle('BZ7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CA')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('CA7', "OOW")->getStyle('CA7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CB')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('CB7', "Total")->getStyle('CB7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('BT5:CB5');
$objPHPExcel->getActiveSheet()->setCellValue('BT5', "New Contract View: 1st year as of $mtcstd")->getStyle('BT5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BU5', "")->getStyle('BU5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BV5', "")->getStyle('BV5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BW5', "")->getStyle('BW5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BX5', "")->getStyle('BX5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BY5', "")->getStyle('BY5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BZ5', "")->getStyle('BZ5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CA5', "")->getStyle('CA5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CB5', "")->getStyle('CB5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CC')->setWidth('1');
$objPHPExcel->getActiveSheet()->setCellValue('CC5', "")->getStyle('CC5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CC6', "")->getStyle('CC6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CC7', "")->getStyle('CC7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('CD6:CF6');
$objPHPExcel->getActiveSheet()->setCellValue('CD6', "Part Number Quantities")->getStyle('CD6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CE6', "")->getStyle('CE6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CF6', "")->getStyle('CF6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CD')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('CD7', "IW")->getStyle('CD7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CE')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('CE7', "OOW")->getStyle('CE7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AR')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('CF7', "Total")->getStyle('CF7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('CG6:CI6');
$objPHPExcel->getActiveSheet()->setCellValue('CG6', "Product Selling Price (PSP) values (".$runc['curr'].")")->getStyle('CG6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CH6', "")->getStyle('CH6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CI6', "")->getStyle('CI6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CG')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('CG7', "IW")->getStyle('CG7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CH')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('CH7', "OOW")->getStyle('CH7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CI')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('CI7', "Total")->getStyle('CI7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('CJ6:CL6');
$objPHPExcel->getActiveSheet()->setCellValue('CJ6', "MQT Unit Quantities")->getStyle('CJ6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CK6', "")->getStyle('CK6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CL6', "")->getStyle('CL6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CJ')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('CJ7', "IW")->getStyle('CJ7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CK')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('CK7', "OOW")->getStyle('CK7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CL')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('CL7', "Total")->getStyle('CL7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('CD5:CL5');
$objPHPExcel->getActiveSheet()->setCellValue('CD5', "New Contract View: 2nd year as of $mtcstd_1")->getStyle('CD5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CE5', "")->getStyle('CE5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CF5', "")->getStyle('CF5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CG5', "")->getStyle('CG5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CH5', "")->getStyle('CH5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CI5', "")->getStyle('CI5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CJ5', "")->getStyle('CJ5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CK5', "")->getStyle('CK5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CL5', "")->getStyle('CL5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('CM6:CO6');
$objPHPExcel->getActiveSheet()->setCellValue('CM6', "Part Number Quantities")->getStyle('CM6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CN6', "")->getStyle('CN6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CO6', "")->getStyle('CO6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CM')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('CM7', "IW")->getStyle('CM7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CN')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('CN7', "OOW")->getStyle('CN7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CO')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('CO7', "Total")->getStyle('CO7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('CP6:CR6');
$objPHPExcel->getActiveSheet()->setCellValue('CP6', "Product Selling Price (PSP) values (".$runc['curr'].")")->getStyle('CP6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CQ6', "")->getStyle('CQ6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CR6', "")->getStyle('CR6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CP')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('CP7', "IW")->getStyle('CP7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CQ')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('CQ7', "OOW")->getStyle('CQ7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CR')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('CR7', "Total")->getStyle('CR7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('CS6:CU6');
$objPHPExcel->getActiveSheet()->setCellValue('CS6', "MQT Unit Quantities")->getStyle('CS6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CT6', "")->getStyle('CT6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CU6', "")->getStyle('CU6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CS')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('CS7', "IW")->getStyle('CS7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CT')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('CT7', "OOW")->getStyle('CT7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('CU')->setWidth('8');
$objPHPExcel->getActiveSheet()->setCellValue('CU7', "Total")->getStyle('CU7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('CM5:CU5');
$objPHPExcel->getActiveSheet()->setCellValue('CM5', "New Contract View: 3rd year as of $mtcstd_2")->getStyle('CM5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CN5', "")->getStyle('CN5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CO5', "")->getStyle('CO5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CP5', "")->getStyle('CP5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CQ5', "")->getStyle('CQ5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CR5', "")->getStyle('CR5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CS5', "")->getStyle('CS5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CT5', "")->getStyle('CT5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('CU5', "")->getStyle('CU5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->freezePane('A8');

$cutrel=Array('TiMOS-C-','TiMOS-B-','TiMOS-MG-C-','TiMOS-I-','TiMOS-DC-B-','TiMOS-DC-C-','1830PSSECX-','1830PSS4-');

foreach ($cutrel as $tmp)
{
	cib_dbQuery("UPDATE `mqt_rep` SET `soft` = REPLACE(`soft`, '$tmp', '')");
}
$no_ts_unit=Array();

while ($ticket16=cib_dbFetchAssoc($res16))
{
	$ntag_n=$ticket16['tag'];
	$ntag_n_d=$ticket16['tagn'];
	
	//$sql5 = "SELECT distinct IF( shtype = '7705-SAR8 v2', '7705-SAR8', shtype) as shtype  FROM `mqt_rep` WHERE 
	//tag='".cib_dbEscape($ticket16['tag'])."' ORDER by shtype ASC";	
	
	$sql5 = "SELECT distinct IF( shtype = '7705-SAR8 v2', '7705-SAR8', shtype) as shtype  
	FROM `mqt_rep` 
	LEFT JOIN `shelf_mqt_conv` AS mqt on `mqt`.`shelf_type_disp`=`mqt_rep`.shtype
	WHERE tag='".cib_dbEscape($ticket16['tag'])."' AND `mqt`.excep='0' ORDER by shtype ASC";
	
	//echo $sql5;
	$res5 = cib_dbQuery($sql5);
	$shelftype='';
	while ($ticket5=cib_dbFetchAssoc($res5))
	{

	if (trim($ticket5['shtype']) =="7705-SAR8")
	{
		$seSelf="shtype like '%".cib_dbEscape($ticket5['shtype'])."%'";
	}
	else $seSelf="shtype='".cib_dbEscape($ticket5['shtype'])."'";

		/*
		$sql8 = "SELECT distinct node FROM `mqt_rep` WHERE 
		tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf ";	
		$nbofnode = cib_dbNumRows(cib_dbQuery($sql8));
		*/

		$sql81 = "SELECT * FROM `shelf_mqt_conv` WHERE `shelf_type_disp`='".cib_dbEscape(trim($ticket5['shtype']))."' LIMIT 1";	
		$mqtname_c = cib_dbFetchAssoc(cib_dbQuery($sql81));		
		
		if ($ticket5['shtype']=="252" OR $ticket5['shtype']=="9471 ATCA") $shelftype="9471 WMM";
		else $shelftype=$ticket5['shtype'];
		$MG='';
		$confi='-';
		if (substr($ticket5['shtype'], -2)=='MG') $MG=' MG';
		
				$sql9_1 = "SELECT * FROM `mqt_cust_profile` WHERE 
				ntag='".cib_dbEscape($ticket16['tag'])."' AND type= '1' AND $seSelf ";			
				$res19_1 = cib_dbFetchAssoc(cib_dbQuery($sql9_1));
				//echo $sql9_1.'<br>';
				$disco=$res19_1['disc'];
				
				//$Res=24+2;
				$Res=0;
				//$Resd=($Res-2);
				$Resd=$Res;
				if (!Empty($res19_1['hw']))
				{
				$Res=$res19_1['hw']+3;
				$Resd=($Res-3);
				}
				
				//$Mqt=12+2;
				$Mqt=0;
				//$Mqtd=($Mqt-2);
				$Mqtd=($Mqt);
				if (!Empty($res19_1['ts']))
				{
				$Mqt=$res19_1['ts']+3;
				$Mqtd=($Mqt-3);
				}
				
				if ($res19_1['confidance']==1) $confi='High';
				else if ($res19_1['confidance']==2) $confi='Medium';				
				
				$sql7 = "SELECT distinct soft FROM `mqt_rep` WHERE 
				tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf AND soft NOT IN ('N/A','')";	
				//echo $sql7;
				$res17 = cib_dbQuery($sql7);

				$shortrelfinal=Array();
				$S_like=0;	
				while ($ticket17=cib_dbFetchAssoc($res17))
				{
					//$relshorta=str_replace($cutrel,'_',$ticket17['soft']);
					//$relshort=explode('.R', $relshorta);
					$relshort=explode('.R', $ticket17['soft']);
					//if (sizeof($relshort) > 1) $shortrelfinal[]=$relshort[0];
					//else $shortrelfinal[]=str_replace($cutrel,'_',$ticket17['soft']);
					//if ($relshort[0]!="N/A") $shortrelfinal[]=$relshort[0];
					isset($relshort[1])? $S_like=1:$S_like=0;					
					$shortrelfinal[]=$relshort[0].':'.$S_like;
					
					}
				
				$shortrelfinal=array_unique($shortrelfinal);			

				foreach ($shortrelfinal as $tmp)
				{	
				$totalamount_units=0;
				$nbofnodes=0;
				$nbofshelves=0;
				$total_nbofboards=0;
				$totaltunitprice=0;	
				
				$tmpa=explode(':', $tmp);
				$S_like=$tmpa[1];
				$tmp=$tmpa[0];
				
					If ($S_like) $S_like_txt="soft like '".cib_dbEscape($tmp)."%'";
					else $S_like_txt="soft = '".cib_dbEscape($tmp)."'";
				
					$reldisp='';
					$reldisp=str_replace($cutrel,'',$tmp);
					$sq200 = "SELECT count(distinct `node`) FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf AND $S_like_txt ";	
					//echo $sq200.'<br>';
					$nbofnodes=cib_dbResult(cib_dbQuery($sq200));
					//$Node_collector_nbofnodes[$ntag_n.':'.$shelftype.':'.$reldisp]=$nbofnodes;				
					
					$sq201 = "SELECT count(*) FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf AND $S_like_txt AND `pn` in ('BR00001AA','IBR00002AA','IBR00003AA','IBR00004AA','IBR00005AA','IBR00006AA','IBR00007AA','IBR00009AA','IBR00010AA') ";	
					//echo $sq200.'<br>';
					$nbofshelves=cib_dbResult(cib_dbQuery($sq201));
					if ($nbofshelves==0) $nbofshelves='-';
					//echo $nbofshelves.'<br>';

					$sql9 = "SELECT distinct 				
					  case 
						when LEFT (`pn`, 3) = '1AB'  THEN  left(`pn`,12) 
						when LEFT (`pn`, 2) = '90' THEN   left(`pn`,9) 
						else  left(`pn`,10) 
					end As pns				
					FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf AND $S_like_txt AND `pn` !='' ";					
					
					$res19 = cib_dbQuery($sql9);				
					$nbofboards='';
					
					while ($ticket19=cib_dbFetchAssoc($res19))
					{										
					$bdescr='';	
					$unitprice=0;
					$tunitprice=0;
					$mqtval=0;
					$mqtvaltotal=0;
					$mqtname='MQT Unit not available';
					$cur='-';
					$sq22 = "SELECT * FROM `tsunits` WHERE 	item='".cib_dbEscape($ticket19['pns'])."' LIMIT 1";	
					//echo $sq22.'<BR>';
					$res22 = cib_dbQuery($sq22);
					$bordch22=cib_dbFetchAssoc($res22);
					$sq23 = "SELECT * FROM `tsunits` WHERE 	item='".substr_replace(cib_dbEscape($ticket19['pns']),'%',8,-1)."' LIMIT 1";	
					//echo $sq23.'<BR>';
					$res23 = cib_dbQuery($sq23);
					$bordch23=cib_dbFetchAssoc($res23);
					$sq24 = "SELECT * FROM `tsunits` WHERE 	item='".substr_replace(cib_dbEscape($ticket19['pns']),'%%',8,2)."' LIMIT 1";	
					//echo $sq24.'<BR>';
					$res24 = cib_dbQuery($sq24);
					$bordch24=cib_dbFetchAssoc($res24);
					//echo $sq24.'<br>';
					if ($bordch22){
					//echo $sq22.'<br>';
					$bdescr=$bordch22['descr'];
					$unitprice=$bordch22['price'];
					$mqtval=$bordch22['valu'];
					$mqtname=$bordch22['unit'];
					$cur=$bordch22['cur'];
					}						
					else if ($bordch23 AND !$bordch22){
					//echo $sq23.'<br>';
					$bdescr=$bordch23['descr'];
					$unitprice=$bordch23['price'];
					$mqtval=$bordch23['valu'];
					$mqtname=$bordch23['unit'];
					$cur=$bordch23['cur'];
					}	
					else if ($bordch24 AND !$bordch22 AND !$bordch23){
					//echo $sq23.'<br>';
					$bdescr=$bordch24['descr'];
					$unitprice=$bordch24['price'];
					$mqtval=$bordch24['valu'];
					$mqtname=$bordch24['unit'];
					$cur=$bordch24['cur'];
					}
					
					$pnslenght=strlen($ticket19['pns']);
					$sq20 = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND $S_like_txt group by `sn`";	
					//echo $sq20."<br>";
					$nbofboards = cib_dbNumRows(cib_dbQuery($sq20));
					
					//echo $total_nbofboards.'<br>';
					
					if (!$bordch24 AND !$bordch22 AND !$bordch23){					
					//$no_ts_unit[]=$ticket19['pns'].':'.$nbofboards;	
					$no_ts_unit[]=$ticket19['pns'];	
					continue;
					}
					$total_nbofboards+=$nbofboards;
					$sq200 = "SELECT count(*) FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND fmd='1' AND $S_like_txt ";	
					//echo $sq200.'<br>';
					$fmdw="";
					$fmd=cib_dbResult(cib_dbQuery($sq200));
					if ($fmd) $fmdw=$fmd." MFG date from common part";
					
					$sql_mtcpst = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Res MONTH) > '$mtcpst_o' AND mandate < '$mtcpst_o' AND $S_like_txt group by `sn`";
					//echo $sql_mtcpst;
					$res_mtcpst = cib_dbQuery($sql_mtcpst);	
					$nb_mtcpst = cib_dbNumRows($res_mtcpst);
					
					$sql_mtcpst_o = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Res MONTH) <= '$mtcpst_o' AND mandate < '$mtcpst_o' AND $S_like_txt group by `sn`";
					
					$sql_mtcpst_o_nmd = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND mandate='0000-00-00' AND $S_like_txt group by `sn`";

					$res_mtcpst_o_nmd = cib_dbQuery($sql_mtcpst_o_nmd);	
					$nb_mtcpst_o_nmd = cib_dbNumRows($res_mtcpst_o_nmd);
					
					$res_mtcpst_o = cib_dbQuery($sql_mtcpst_o);	
					$nb_mtcpst_o = cib_dbNumRows($res_mtcpst_o);					
					
					$sql_mtcstd = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Res MONTH) > '$mtcstd_o' AND mandate < '$mtcstd_o' AND $S_like_txt group by `sn`";
					$res_mtcstd = cib_dbQuery($sql_mtcstd);	
					$nb_mtcstd = cib_dbNumRows($res_mtcstd);
					
					$sql_mtcstd_o = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Res MONTH) <= '$mtcstd_o' AND mandate < '$mtcstd_o' AND $S_like_txt group by `sn`";
					
					$sql_mtcstd_o_nmd = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND mandate='0000-00-00' AND $S_like_txt group by `sn`";

					$res_mtcstd_o_nmd = cib_dbQuery($sql_mtcstd_o_nmd);	
					$nb_mtcstd_o_nmd = cib_dbNumRows($res_mtcstd_o_nmd);					
					
					$res_mtcstd_o = cib_dbQuery($sql_mtcstd_o);	
					$nb_mtcstd_o = cib_dbNumRows($res_mtcstd_o);	

					$sql_mtcstd1 = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Res MONTH) > '$mtcstd_o1' AND mandate < '$mtcstd_o1' AND $S_like_txt group by `sn`";
					$res_mtcstd1 = cib_dbQuery($sql_mtcstd1);	
					$nb_mtcstd1 = cib_dbNumRows($res_mtcstd1);
					
					$sql_mtcstd1_o = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Res MONTH) <= '$mtcstd_o1' AND mandate < '$mtcstd_o1' AND $S_like_txt group by `sn`";
					
					$sql_mtcstd1_o_nmd = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND mandate='0000-00-00' AND $S_like_txt group by `sn`";
					
					$res_mtcstd1_o_nmd = cib_dbQuery($sql_mtcstd1_o_nmd);	
					$nb_mtcstd1_o_nmd = cib_dbNumRows($res_mtcstd1_o_nmd);					
					
					$res_mtcstd1_o = cib_dbQuery($sql_mtcstd1_o);	
					$nb_mtcstd1_o = cib_dbNumRows($res_mtcstd1_o);					
					
					$sql_mtcstd2 = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Res MONTH) > '$mtcstd_o2' AND mandate < '$mtcstd_o2' AND $S_like_txt group by `sn`";
					$res_mtcstd2 = cib_dbQuery($sql_mtcstd2);	
					$nb_mtcstd2 = cib_dbNumRows($res_mtcstd2);
					
					$sql_mtcstd2_o = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Res MONTH) <= '$mtcstd_o2' AND mandate < '$mtcstd_o2' AND $S_like_txt group by `sn`";
					
					$sql_mtcstd2_o_nmd = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND mandate='0000-00-00' AND $S_like_txt group by `sn`";
					
					$res_mtcstd2_o_nmd = cib_dbQuery($sql_mtcstd2_o_nmd);	
					$nb_mtcstd2_o_nmd = cib_dbNumRows($res_mtcstd2_o_nmd);
					
					$res_mtcstd2_o = cib_dbQuery($sql_mtcstd2_o);	
					$nb_mtcstd2_o = cib_dbNumRows($res_mtcstd2_o);					
		
					if ($ticket19['pns']=="N/A") $partnumber="Unknown";		 
					else $partnumber=$ticket19['pns'];						
					
					if ($cur=='USD' AND $runc['curr']=='EUR') $unitprice=$unitprice*$USD_EUR;
					else if ($cur=='EUR' AND $runc['curr']=='USD') $unitprice=$unitprice*$EUR_USD;					
					
					$tunitprice=$nbofboards*$unitprice;
					//$mqtvaltotal=round($nbofboards*$mqtval);
					$mqtvaltotal=$nbofboards*$mqtval;
					$totalamount_units+=$mqtvaltotal;
					$totaltunitprice+=$tunitprice;
					
					$sql_mtcpst_mqt = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Mqt MONTH) > '$mtcpst_o' AND mandate < '$mtcpst_o' AND $S_like_txt group by `sn`";
					$res_mtcpst_mqt = cib_dbQuery($sql_mtcpst_mqt);	
					$nb_mtcpst_mqt = cib_dbNumRows($res_mtcpst_mqt);
					$mqt_iw=($nb_mtcpst_mqt*$mqtval);
					
					$sql_mtcpst_mqt_o = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Mqt MONTH) <= '$mtcpst_o' AND mandate < '$mtcpst_o' AND $S_like_txt group by `sn`";
					
					$sql_mtcpst_mqt_o_nmd = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND mandate='0000-00-00' AND $S_like_txt group by `sn`";
					
					//$sql_mtcpst_mqt_o = "SELECT id FROM `mqt_rep` WHERE 
					//tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND 
					//if(mandate='0000-00-00',mandate='0000-00-00', ADDDATE(mandate, INTERVAL $Mqt MONTH) <= '$mtcpst_o') AND mandate < '$mtcpst_o' AND $S_like_txt group by `sn`";
					
					$res_mtcpst_mqt_o = cib_dbQuery($sql_mtcpst_mqt_o);	
					$nb_mtcpst_mqt_o = cib_dbNumRows($res_mtcpst_mqt_o);
					
					$res_mtcpst_mqt_o_nmd = cib_dbQuery($sql_mtcpst_mqt_o_nmd);	
					$nb_mtcpst_mqt_o_nmd = cib_dbNumRows($res_mtcpst_mqt_o_nmd);
					
					$mqt_iw_o=($nb_mtcpst_mqt_o*$mqtval);

					$sql_mtcstd_mqt = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Mqt MONTH) > '$mtcstd_o' AND mandate < '$mtcstd_o' AND $S_like_txt group by `sn`";
					$res_mtcstd_mqt = cib_dbQuery($sql_mtcstd_mqt);	
					$nb_mtcstd_mqt = cib_dbNumRows($res_mtcstd_mqt);
					$mqt_mtn_iw=($nb_mtcstd_mqt*$mqtval);
	
					$sql_mtcstd_mqt_o = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Mqt MONTH) <= '$mtcstd_o' AND mandate < '$mtcstd_o' AND $S_like_txt group by `sn`";
					
					$sql_mtcstd_mqt_o_nmd = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND mandate='0000-00-00'  AND $S_like_txt group by `sn`";

					$res_mtcstd_mqt_o_nmd = cib_dbQuery($sql_mtcstd_mqt_o_nmd);	
					$nb_mtcstd_mqt_o_nmd = cib_dbNumRows($res_mtcstd_mqt_o_nmd);					
					
					$res_mtcstd_mqt_o = cib_dbQuery($sql_mtcstd_mqt_o);	
					$nb_mtcstd_mqt_o = cib_dbNumRows($res_mtcstd_mqt_o);
					
					$mqt_mtn_iw_o=($nb_mtcstd_mqt_o*$mqtval);
					
					$sql_mtcstd1_mqt = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Mqt MONTH) > '$mtcstd_o1' AND mandate < '$mtcstd_o1' AND $S_like_txt group by `sn`";
					$res_mtcstd1_mqt = cib_dbQuery($sql_mtcstd1_mqt);	
					$nb_mtcstd1_mqt = cib_dbNumRows($res_mtcstd1_mqt);
					$mqt_mtn_iw1=($nb_mtcstd1_mqt*$mqtval);

	
					$sql_mtcstd_mqt1_o = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Mqt MONTH) <= '$mtcstd_o1' AND mandate < '$mtcstd_o1' AND $S_like_txt group by `sn`";
					
					$sql_mtcstd_mqt1_o_nmd = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND mandate='0000-00-00' AND $S_like_txt group by `sn`";
					
					$res_mtcstd_mqt1_o_nmd = cib_dbQuery($sql_mtcstd_mqt1_o_nmd);						
					$nb_mtcstd_mqt1_o_nmd = cib_dbNumRows($res_mtcstd_mqt1_o_nmd);

					$res_mtcstd_mqt1_o = cib_dbQuery($sql_mtcstd_mqt1_o);						
					$nb_mtcstd_mqt1_o = cib_dbNumRows($res_mtcstd_mqt1_o);
					
					$mqt_mtn_iw1_o=($nb_mtcstd_mqt1_o*$mqtval);				
					

					$sql_mtcstd2_mqt = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Mqt MONTH) > '$mtcstd_o2' AND mandate < '$mtcstd_o2' AND $S_like_txt group by `sn`";
					$res_mtcstd2_mqt = cib_dbQuery($sql_mtcstd2_mqt);	
					$nb_mtcstd2_mqt = cib_dbNumRows($res_mtcstd2_mqt);
					$mqt_mtn_iw2=($nb_mtcstd2_mqt*$mqtval);

	
					$sql_mtcstd_mqt2_o = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND ADDDATE(mandate, INTERVAL $Mqt MONTH) <= '$mtcstd_o2' AND mandate < '$mtcstd_o2' AND $S_like_txt group by `sn`";
					
					
					$sql_mtcstd_mqt2_o_nmd = "SELECT id FROM `mqt_rep` WHERE 
					tag='".cib_dbEscape($ticket16['tag'])."' AND $seSelf  AND left( `pn` , $pnslenght )='".cib_dbEscape($ticket19['pns'])."' AND mandate='0000-00-00'  AND $S_like_txt group by `sn`";
					
					$res_mtcstd_mqt2_o_nmd = cib_dbQuery($sql_mtcstd_mqt2_o_nmd);					
					$nb_mtcstd_mqt2_o_nmd = cib_dbNumRows($res_mtcstd_mqt2_o_nmd);
					
					$res_mtcstd_mqt2_o = cib_dbQuery($sql_mtcstd_mqt2_o);					
					$nb_mtcstd_mqt2_o = cib_dbNumRows($res_mtcstd_mqt2_o);
					
					$mqt_mtn_iw2_o=($nb_mtcstd_mqt2_o*$mqtval);				
					
					$warn='';
					$Mqtdy_t=yearfm($Mqtd);
					$Resd_t=yearfm($Resd);
					//$reldisp=str_replace($cutrel,'',$tmp);
					if ($nb_mtcpst_mqt_o_nmd >0) {
					$warn='No Manufactured date';
					$Resd_t='Unknown';
					$Mqtdy_t='Unknown';
					}
					
					$mqt_product = (!Empty ($mqtname_c['mqt_product']) ? $mqtname_c['mqt_product'] : 'MQT details not available');
					$mqt_config = (!Empty ($mqtname_c['mqt_config']) ? $mqtname_c['mqt_config'] : 'MQT details not available');
					$pr_com_name = (!Empty ($mqtname_c['pr_com_name']) ? $mqtname_c['pr_com_name'] : 'MQT details not available');
					$mqt_u_name = (!Empty ($mqtname_c['mqt_u_name']) ? $mqtname_c['mqt_u_name'] : 'MQT details not available');
					
					//TAB 2 collect//
					$MQT_Prod[]=$mqt_product.':'.$mqt_config.':'.$reldisp;
					$MQT_Qty[$mqt_product.':'.$mqt_config.':'.$reldisp][]=$nbofboards;
					//$MQT_PSP[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($unitprice-($unitprice*($res19_1['disc']/100)))*$nbofboards;
					$MQT_OPD[$mqt_product.':'.$mqt_config.':'.$reldisp][]=$res19_1['disc'];
					//$MQT_Uname[$mqt_product.':'.$mqt_config.':'.$reldisp]=$mqtname;
					$MQT_Uname[$mqt_product.':'.$mqt_config.':'.$reldisp]=$mqtname_c['mqt_u_name'];
					$MQT_AD[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcpst_mqt+$nb_mtcpst_mqt_o+$nb_mtcpst_mqt_o_nmd)*$mqtval;
					$MQT_AA[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcpst_mqt+$nb_mtcpst_mqt_o+$nb_mtcpst_mqt_o_nmd)*($unitprice-($unitprice*($res19_1['disc']/100)));
					$MQT_O[$mqt_product.':'.$mqt_config.':'.$reldisp]=$confi;
					$MQT_AB[$mqt_product.':'.$mqt_config.':'.$reldisp][]=$nb_mtcpst_mqt*$mqtval;
					$MQT_BP[$mqt_product.':'.$mqt_config.':'.$reldisp][]=$nb_mtcpst*$mqtval;
					$MQT_AC[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcpst_mqt_o+$nb_mtcpst_mqt_o_nmd)*$mqtval;
					//$MQT_Z[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcpst_mqt_o+$nb_mtcpst_mqt_o_nmd)*($unitprice-($unitprice*($res19_1['disc']/100)));
					$MQT_BQ[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcpst_o+$nb_mtcpst_o_nmd)*$mqtval;
					$MQT_AA_FP[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcpst_mqt+$nb_mtcpst_mqt_o+$nb_mtcpst_mqt_o_nmd)*$unitprice;
					$MQT_Z_FP[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcpst_mqt_o+$nb_mtcpst_mqt_o_nmd)*$unitprice;
					$MQT_BN_FP[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcpst_o+$nb_mtcpst_o_nmd)*$unitprice;
					$MQT_AN[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcstd_mqt+$nb_mtcstd_mqt_o+$nb_mtcstd_mqt_o_nmd)*$mqtval;
					$MQT_AK[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcstd_mqt+$nb_mtcstd_mqt_o+$nb_mtcstd_mqt_o_nmd)*($unitprice-($unitprice*($res19_1['disc']/100)));
					$MQT_AK_FP[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcstd_mqt+$nb_mtcstd_mqt_o+$nb_mtcstd_mqt_o_nmd)*$unitprice;
					$MQT_AL[$mqt_product.':'.$mqt_config.':'.$reldisp][]=$nb_mtcstd_mqt*$mqtval;
					$MQT_BZ[$mqt_product.':'.$mqt_config.':'.$reldisp][]=$nb_mtcstd*$mqtval;
					$MQT_AN2[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcstd_mqt_o+$nb_mtcstd_mqt_o_nmd)*$mqtval;
					$MQT_AO_FP[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcstd_mqt_o+$nb_mtcstd_mqt_o_nmd)*$unitprice;
					$MQT_AO[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcstd_mqt_o+$nb_mtcstd_mqt_o_nmd)*($unitprice-($unitprice*($res19_1['disc']/100)));
					$MQT_AQ[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcstd_o+$nb_mtcstd_o_nmd)*$mqtval;
					$MQT_AR_FP[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcstd_o+$nb_mtcstd_o_nmd)*$unitprice;
					$MQT_AR[$mqt_product.':'.$mqt_config.':'.$reldisp][]=($nb_mtcstd_o+$nb_mtcstd_o_nmd)*($unitprice-($unitprice*($res19_1['disc']/100)));
					$objPHPExcel->getActiveSheet()->setCellValue('B'.$i, $cnt)->getStyle('B'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('G'.$i, $ticket16['tagn'])->getStyle('G'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('H'.$i, $warn.$fmdw)->getStyle('H'.$i)->applyFromArray($styleArray01);		
					$objPHPExcel->getActiveSheet()->setCellValue('I'.$i, $mqt_product)->getStyle('I'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('J'.$i, $mqt_config)->getStyle('J'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('K'.$i, '\''.$reldisp)->getStyle('K'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->getStyle('K'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_TEXT);					
					$objPHPExcel->getActiveSheet()->setCellValue('C'.$i, $partnumber)->getStyle('C'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('D'.$i, $bdescr)->getStyle('D'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('E'.$i, $shelftype)->getStyle('E'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('F'.$i, $nbofboards)->getStyle('F'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('L'.$i, $unitprice)->getStyle('L'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->getStyle('L'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);
					$objPHPExcel->getActiveSheet()->setCellValue('M'.$i, '=L'.$i.'*F'.$i.'')->getStyle('M'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->getStyle('M'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);
					$objPHPExcel->getActiveSheet()->getStyle('N'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);					
					$objPHPExcel->getActiveSheet()->setCellValue('N'.$i, '='.$res19_1['disc']/100)->getStyle('N'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('O'.$i, $confi)->getStyle('O'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('P'.$i, '=L'.$i.'-(L'.$i.'*N'.$i.')')->getStyle('P'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('Q'.$i, '=P'.$i.'*F'.$i.'')->getStyle('Q'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->getStyle('Q'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);
					$objPHPExcel->getActiveSheet()->setCellValue('R'.$i, $mqtval)->getStyle('R'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('S'.$i, '=R'.$i.'*F'.$i.'')->getStyle('S'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('T'.$i, $mqtname)->getStyle('T'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('U'.$i, ' '.$Mqtdy_t)->getStyle('U'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('V'.$i, $nb_mtcpst_mqt)->getStyle('V'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('W'.$i, '='.$nb_mtcpst_mqt_o.'+'.$nb_mtcpst_mqt_o_nmd.'')->getStyle('W'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('X'.$i, '=V'.$i.'+W'.$i.'')->getStyle('X'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('Y'.$i, '=V'.$i.'*P'.$i.'')->getStyle('Y'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('Z'.$i, '=W'.$i.'*P'.$i.'')->getStyle('Z'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('AA'.$i, '=Y'.$i.'+Z'.$i.'')->getStyle('AA'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->getStyle('AA'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);					
					//$objPHPExcel->getActiveSheet()->setCellValue('AB'.$i, $mqt_iw)->getStyle('AB'.$i)->applyFromArray($styleArray01);					
					//$objPHPExcel->getActiveSheet()->setCellValue('AC'.$i, $mqt_iw_o)->getStyle('AC'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('AB'.$i, '=V'.$i.'*R'.$i.'')->getStyle('AB'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('AC'.$i, '=W'.$i.'*R'.$i.'')->getStyle('AC'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('AD'.$i, '=AB'.$i.'+AC'.$i.'')->getStyle('AD'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('AE'.$i, $mqtname)->getStyle('AE'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('AF'.$i, $nb_mtcstd_mqt)->getStyle('AF'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('AG'.$i, '='.$nb_mtcstd_mqt_o.'+'.$nb_mtcstd_mqt_o_nmd.'')->getStyle('AG'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('AH'.$i, '=AF'.$i.'+AG'.$i.'')->getStyle('AH'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('AI'.$i, '=AF'.$i.'*P'.$i.'')->getStyle('AI'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('AJ'.$i, '=AG'.$i.'*P'.$i.'')->getStyle('AJ'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('AK'.$i, '=AI'.$i.'+AJ'.$i.'')->getStyle('AK'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->getStyle('AK'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);					
					//$objPHPExcel->getActiveSheet()->setCellValue('AL'.$i, $mqt_mtn_iw)->getStyle('AL'.$i)->applyFromArray($styleArray01);					
					//$objPHPExcel->getActiveSheet()->setCellValue('AM'.$i, $mqt_mtn_iw_o)->getStyle('AM'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('AL'.$i, '=AF'.$i.'*R'.$i.'')->getStyle('AL'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('AM'.$i, '=AG'.$i.'*R'.$i.'')->getStyle('AM'.$i)->applyFromArray($styleArray01);	
					$objPHPExcel->getActiveSheet()->setCellValue('AN'.$i, '=AL'.$i.'+AM'.$i.'')->getStyle('AN'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('AO'.$i, '')->getStyle('AO'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('AP'.$i, $nb_mtcstd1_mqt)->getStyle('AP'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('AQ'.$i, '='.$nb_mtcstd_mqt1_o.'+'.$nb_mtcstd_mqt1_o_nmd.'')->getStyle('AQ'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('AR'.$i, '=AP'.$i.'+AQ'.$i.'')->getStyle('AR'.$i)->applyFromArray($styleArray01);	
					$objPHPExcel->getActiveSheet()->setCellValue('AS'.$i, '=AP'.$i.'*P'.$i.'')->getStyle('AS'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('AT'.$i, '=AQ'.$i.'*P'.$i.'')->getStyle('AT'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('AU'.$i, '=AS'.$i.'+AT'.$i.'')->getStyle('AU'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->getStyle('AU'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);
					//$objPHPExcel->getActiveSheet()->setCellValue('AV'.$i, $mqt_mtn_iw1)->getStyle('AV'.$i)->applyFromArray($styleArray01);					
					//$objPHPExcel->getActiveSheet()->setCellValue('AW'.$i, $mqt_mtn_iw1_o)->getStyle('AW'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('AV'.$i, '=AP'.$i.'*R'.$i.'')->getStyle('AV'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('AW'.$i, '=AQ'.$i.'*R'.$i.'')->getStyle('AW'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('AX'.$i, '=AV'.$i.'+AW'.$i.'')->getStyle('AX'.$i)->applyFromArray($styleArray01);	
					$objPHPExcel->getActiveSheet()->setCellValue('AY'.$i, $nb_mtcstd2_mqt)->getStyle('AY'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('AZ'.$i, '='.$nb_mtcstd_mqt2_o.'+'.$nb_mtcstd_mqt2_o_nmd.'')->getStyle('AZ'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('BA'.$i, '=AY'.$i.'+AZ'.$i.'')->getStyle('BA'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('BB'.$i, '=AY'.$i.'*P'.$i.'')->getStyle('BB'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('BC'.$i, '=AZ'.$i.'*P'.$i.'')->getStyle('BC'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('BD'.$i, '=BB'.$i.'+BC'.$i.'')->getStyle('BD'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->getStyle('BD'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);
					//$objPHPExcel->getActiveSheet()->setCellValue('BE'.$i, $mqt_mtn_iw2)->getStyle('BE'.$i)->applyFromArray($styleArray01);					
					//$objPHPExcel->getActiveSheet()->setCellValue('BF'.$i, $mqt_mtn_iw2_o)->getStyle('BF'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('BE'.$i, '=AY'.$i.'*R'.$i.'')->getStyle('BE'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('BF'.$i, '=AZ'.$i.'*R'.$i.'')->getStyle('BF'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('BG'.$i, '=BE'.$i.'+BF'.$i.'')->getStyle('BG'.$i)->applyFromArray($styleArray01);	
					$objPHPExcel->getActiveSheet()->setCellValue('BH'.$i, "")->getStyle('BH'.$i)->applyFromArray($styleArray03);
					$objPHPExcel->getActiveSheet()->setCellValue('BI'.$i, ' '.$Resd_t)->getStyle('BI'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('BJ'.$i, $nb_mtcpst)->getStyle('BJ'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('BK'.$i, '='.$nb_mtcpst_o.'+'.$nb_mtcpst_o_nmd.'')->getStyle('BK'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('BL'.$i, '=BJ'.$i.'+BK'.$i.'')->getStyle('BL'.$i)->applyFromArray($styleArray01);	
					$objPHPExcel->getActiveSheet()->setCellValue('BM'.$i, '=BJ'.$i.'*P'.$i.'')->getStyle('BM'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('BN'.$i, '=BK'.$i.'*P'.$i.'')->getStyle('BN'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('BO'.$i, '=BM'.$i.'+BN'.$i.'')->getStyle('BO'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->getStyle('BO'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);					
					$objPHPExcel->getActiveSheet()->setCellValue('BP'.$i, '=BJ'.$i.'*R'.$i.'')->getStyle('BP'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('BQ'.$i, '=BK'.$i.'*R'.$i.'')->getStyle('BQ'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('BR'.$i, '=BP'.$i.'+BQ'.$i.'')->getStyle('BR'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('BS'.$i, $mqtname)->getStyle('BS'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('BT'.$i, $nb_mtcstd)->getStyle('BT'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('BU'.$i, '='.$nb_mtcstd_o.'+'.$nb_mtcstd_o_nmd.'')->getStyle('BU'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('BV'.$i, '=BT'.$i.'+BU'.$i.'')->getStyle('BV'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('BW'.$i, '=BT'.$i.'*P'.$i.'')->getStyle('BW'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('BX'.$i, '=BU'.$i.'*P'.$i.'')->getStyle('BX'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('BY'.$i, '=BW'.$i.'+BX'.$i.'')->getStyle('BY'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->getStyle('BY'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);		
					$objPHPExcel->getActiveSheet()->setCellValue('BZ'.$i, '=BT'.$i.'*R'.$i.'')->getStyle('BZ'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('CA'.$i, '=BU'.$i.'*R'.$i.'')->getStyle('CA'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('CB'.$i, '=BZ'.$i.'+CA'.$i.'')->getStyle('CB'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('CC'.$i, '')->getStyle('CC'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('CD'.$i, $nb_mtcstd1)->getStyle('CD'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('CE'.$i, '='.$nb_mtcstd1_o.'+'.$nb_mtcstd1_o_nmd.'')->getStyle('CE'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('CF'.$i, '=CD'.$i.'+CE'.$i.'')->getStyle('CF'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('CG'.$i, '=CD'.$i.'*P'.$i.'')->getStyle('CG'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('CH'.$i, '=CE'.$i.'*P'.$i.'')->getStyle('CH'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('CI'.$i, '=CG'.$i.'+CH'.$i.'')->getStyle('CI'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->getStyle('CI'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);
					$objPHPExcel->getActiveSheet()->setCellValue('CJ'.$i, '=CD'.$i.'*R'.$i.'')->getStyle('CJ'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('CK'.$i, '=CE'.$i.'*R'.$i.'')->getStyle('CK'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('CL'.$i, '=CJ'.$i.'+CA'.$i.'')->getStyle('CL'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('CM'.$i, $nb_mtcstd2)->getStyle('CM'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('CN'.$i, '='.$nb_mtcstd2_o.'+'.$nb_mtcstd2_o_nmd.'')->getStyle('CN'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('CO'.$i, '=CM'.$i.'+CN'.$i.'')->getStyle('CO'.$i)->applyFromArray($styleArray01);	
					$objPHPExcel->getActiveSheet()->setCellValue('CP'.$i, '=CM'.$i.'*P'.$i.'')->getStyle('CP'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('CQ'.$i, '=CN'.$i.'*P'.$i.'')->getStyle('CQ'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('CR'.$i, '=CP'.$i.'+CQ'.$i.'')->getStyle('CR'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->getStyle('CR'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);					
					$objPHPExcel->getActiveSheet()->setCellValue('CS'.$i, '=CM'.$i.'*R'.$i.'')->getStyle('CS'.$i)->applyFromArray($styleArray01);					
					$objPHPExcel->getActiveSheet()->setCellValue('CT'.$i, '=CN'.$i.'*R'.$i.'')->getStyle('CT'.$i)->applyFromArray($styleArray01);
					$objPHPExcel->getActiveSheet()->setCellValue('CU'.$i, '=CS'.$i.'+CT'.$i.'')->getStyle('CU'.$i)->applyFromArray($styleArray01);
					$i++;
					$cnt++;
					}
				$Node_collector[]=$ntag_n_d.':'.$ntag_n.':'.$shelftype.':'.$mqt_product.':'.$mqt_config.':'.$confi.':'.$disco.':'.$reldisp.':'.$nbofnodes.':'.$totalamount_units.':'.$total_nbofboards.':'.$pr_com_name.':'.$mqt_u_name.':'.$totaltunitprice.':'.$nbofshelves;
				//if ($shelftype!="Pluggable") $Node_collector[]=$ntag_n_d.':'.$ntag_n.':'.$shelftype.':'.$mqt_product.':'.$mqt_config.':'.$confi.':'.$disco.':'.$reldisp.':'.$nbofnodes.':'.$totalamount_units.':'.$total_nbofboards.':'.$pr_com_name.':'.$mqt_u_name.':'.$totaltunitprice.':'.$nbofshelves;
				
				}	
//$disco=$res19_1['disc'];
//echo $total_nbofboards.'<br>';
//echo $totalamount_units.'<br>';
//$Node_collector[]=$ntag_n_d.':'.$ntag_n.':'.$shelftype.':'.$mqt_product.':'.$mqt_config.':'.$confi.':'.$disco.':'.$reldisp.':'.$nbofnodes.':'.$totalamount_units.':'.$total_nbofboards;				
//echo $totalamount_units.'<br>';					
	}

}

$cnt=$cnt+7;
$objPHPExcel->getActiveSheet()->setAutoFilter('B7:U7');
//$objPHPExcel->getActiveSheet()->setAutoFilter('AE7:AE7');
//$objPHPExcel->getActiveSheet()->setCellValue('N'.$cnt, '=SUBTOTAL(9,N8:N'.($cnt-1).')')->getStyle('N'.$cnt)->applyFromArray($styleArray01);

If ($runc['curr']=='EUR') 
{
	$objPHPExcel->getActiveSheet()->getStyle('M'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	//$objPHPExcel->getActiveSheet()->getStyle('M'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
	$objPHPExcel->getActiveSheet()->getStyle('Q'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AA'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('Y'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('Z'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AI'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AJ'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AK'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AS'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AT'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AU'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('BB'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('BC'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('BD'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('BM'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('BN'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('BO'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('BW'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('BX'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('BY'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('CG'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('CH'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('CI'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('CP'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('CQ'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('CR'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
}
else 
{
	$objPHPExcel->getActiveSheet()->getStyle('M'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);	
	$objPHPExcel->getActiveSheet()->getStyle('Q'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AA'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('Y'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('Z'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AI'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AJ'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AK'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AS'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AT'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AU'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('BB'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('BC'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('BD'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('BM'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('BN'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('BO'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('BW'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('BX'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('BY'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('CH'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('CH'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('CI'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('CP'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('CQ'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('CR'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
}

$objPHPExcel->getActiveSheet()->setCellValue('M'.$cnt, '=SUBTOTAL(9,M8:M'.($cnt-1).')')->getStyle('M'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AA'.$cnt, '=SUBTOTAL(9,AA8:AA'.($cnt-1).')')->getStyle('AA'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('Q'.$cnt, '=SUBTOTAL(9,Q8:Q'.($cnt-1).')')->getStyle('Q'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('S'.$cnt, '=SUBTOTAL(9,S8:S'.($cnt-1).')')->getStyle('S'.$cnt)->applyFromArray($styleArray01);		
$objPHPExcel->getActiveSheet()->setCellValue('V'.$cnt, '=SUBTOTAL(9,V8:V'.($cnt-1).')')->getStyle('V'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('W'.$cnt, '=SUBTOTAL(9,W8:W'.($cnt-1).')')->getStyle('W'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('X'.$cnt, '=SUBTOTAL(9,X8:X'.($cnt-1).')')->getStyle('X'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('Y'.$cnt, '=SUBTOTAL(9,Y8:Y'.($cnt-1).')')->getStyle('Y'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('Z'.$cnt, '=SUBTOTAL(9,Z8:Z'.($cnt-1).')')->getStyle('Z'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AB'.$cnt, '=SUBTOTAL(9,AB8:AB'.($cnt-1).')')->getStyle('AB'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AC'.$cnt, '=SUBTOTAL(9,AC8:AC'.($cnt-1).')')->getStyle('AC'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AD'.$cnt, '=SUBTOTAL(9,AD8:AD'.($cnt-1).')')->getStyle('AD'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('AF'.$cnt, '=SUBTOTAL(9,AF8:AF'.($cnt-1).')')->getStyle('AF'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AG'.$cnt, '=SUBTOTAL(9,AG8:AG'.($cnt-1).')')->getStyle('AG'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AH'.$cnt, '=SUBTOTAL(9,AH8:AH'.($cnt-1).')')->getStyle('AH'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AI'.$cnt, '=SUBTOTAL(9,AI8:AI'.($cnt-1).')')->getStyle('AI'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AJ'.$cnt, '=SUBTOTAL(9,AJ8:AJ'.($cnt-1).')')->getStyle('AJ'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AK'.$cnt, '=SUBTOTAL(9,AK8:AK'.($cnt-1).')')->getStyle('AK'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AL'.$cnt, '=SUBTOTAL(9,AL8:AL'.($cnt-1).')')->getStyle('AL'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AM'.$cnt, '=SUBTOTAL(9,AM8:AM'.($cnt-1).')')->getStyle('AM'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AN'.$cnt, '=SUBTOTAL(9,AN8:AN'.($cnt-1).')')->getStyle('AN'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('AP'.$cnt, '=SUBTOTAL(9,AP8:AP'.($cnt-1).')')->getStyle('AP'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('AQ'.$cnt, '=SUBTOTAL(9,AQ8:AQ'.($cnt-1).')')->getStyle('AQ'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('AR'.$cnt, '=SUBTOTAL(9,AR8:AR'.($cnt-1).')')->getStyle('AR'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('AS'.$cnt, '=SUBTOTAL(9,AS8:AS'.($cnt-1).')')->getStyle('AS'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('AT'.$cnt, '=SUBTOTAL(9,AT8:AT'.($cnt-1).')')->getStyle('AT'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('AU'.$cnt, '=SUBTOTAL(9,AU8:AU'.($cnt-1).')')->getStyle('AU'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('AV'.$cnt, '=SUBTOTAL(9,AV8:AV'.($cnt-1).')')->getStyle('AV'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('AW'.$cnt, '=SUBTOTAL(9,AW8:AW'.($cnt-1).')')->getStyle('AW'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('AX'.$cnt, '=SUBTOTAL(9,AX8:AX'.($cnt-1).')')->getStyle('AX'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('AY'.$cnt, '=SUBTOTAL(9,AY8:AY'.($cnt-1).')')->getStyle('AY'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('AZ'.$cnt, '=SUBTOTAL(9,AZ8:AZ'.($cnt-1).')')->getStyle('AZ'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BA'.$cnt, '=SUBTOTAL(9,BA8:BA'.($cnt-1).')')->getStyle('BA'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BB'.$cnt, '=SUBTOTAL(9,BB8:BB'.($cnt-1).')')->getStyle('BB'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BC'.$cnt, '=SUBTOTAL(9,BC8:BC'.($cnt-1).')')->getStyle('BC'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BD'.$cnt, '=SUBTOTAL(9,BD8:BD'.($cnt-1).')')->getStyle('BD'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BE'.$cnt, '=SUBTOTAL(9,BE8:BE'.($cnt-1).')')->getStyle('BE'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BF'.$cnt, '=SUBTOTAL(9,BF8:BF'.($cnt-1).')')->getStyle('BF'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BG'.$cnt, '=SUBTOTAL(9,BG8:BG'.($cnt-1).')')->getStyle('BG'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('BJ'.$cnt, '=SUBTOTAL(9,BJ8:BJ'.($cnt-1).')')->getStyle('BJ'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BK'.$cnt, '=SUBTOTAL(9,BK8:BK'.($cnt-1).')')->getStyle('BK'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BL'.$cnt, '=SUBTOTAL(9,BL8:BL'.($cnt-1).')')->getStyle('BL'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BM'.$cnt, '=SUBTOTAL(9,BM8:BM'.($cnt-1).')')->getStyle('BM'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BN'.$cnt, '=SUBTOTAL(9,BN8:BN'.($cnt-1).')')->getStyle('BN'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BO'.$cnt, '=SUBTOTAL(9,BO8:BO'.($cnt-1).')')->getStyle('BO'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BP'.$cnt, '=SUBTOTAL(9,BP8:BP'.($cnt-1).')')->getStyle('BP'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BQ'.$cnt, '=SUBTOTAL(9,BQ8:BQ'.($cnt-1).')')->getStyle('BQ'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BR'.$cnt, '=SUBTOTAL(9,BR8:BR'.($cnt-1).')')->getStyle('BR'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('BT'.$cnt, '=SUBTOTAL(9,BT8:BT'.($cnt-1).')')->getStyle('BT'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('BU'.$cnt, '=SUBTOTAL(9,BU8:BU'.($cnt-1).')')->getStyle('BU'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('BV'.$cnt, '=SUBTOTAL(9,BV8:BV'.($cnt-1).')')->getStyle('BV'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('BW'.$cnt, '=SUBTOTAL(9,BW8:BW'.($cnt-1).')')->getStyle('BW'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('BX'.$cnt, '=SUBTOTAL(9,BX8:BX'.($cnt-1).')')->getStyle('BX'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('BY'.$cnt, '=SUBTOTAL(9,BY8:BY'.($cnt-1).')')->getStyle('BY'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('BZ'.$cnt, '=SUBTOTAL(9,BZ8:BZ'.($cnt-1).')')->getStyle('BZ'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('CA'.$cnt, '=SUBTOTAL(9,CA8:CA'.($cnt-1).')')->getStyle('CA'.$cnt)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('CB'.$cnt, '=SUBTOTAL(9,CB8:CB'.($cnt-1).')')->getStyle('CB'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CD'.$cnt, '=SUBTOTAL(9,CD8:CD'.($cnt-1).')')->getStyle('CD'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CE'.$cnt, '=SUBTOTAL(9,CE8:CE'.($cnt-1).')')->getStyle('CE'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CF'.$cnt, '=SUBTOTAL(9,CF8:CF'.($cnt-1).')')->getStyle('CF'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CG'.$cnt, '=SUBTOTAL(9,CG8:CG'.($cnt-1).')')->getStyle('CG'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CH'.$cnt, '=SUBTOTAL(9,CH8:CH'.($cnt-1).')')->getStyle('CH'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CI'.$cnt, '=SUBTOTAL(9,Ci8:CI'.($cnt-1).')')->getStyle('CI'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CJ'.$cnt, '=SUBTOTAL(9,CJ8:CJ'.($cnt-1).')')->getStyle('CJ'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CK'.$cnt, '=SUBTOTAL(9,CK8:CK'.($cnt-1).')')->getStyle('CK'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CL'.$cnt, '=SUBTOTAL(9,CL8:CL'.($cnt-1).')')->getStyle('CL'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CM'.$cnt, '=SUBTOTAL(9,CM8:CM'.($cnt-1).')')->getStyle('CM'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CN'.$cnt, '=SUBTOTAL(9,CN8:CN'.($cnt-1).')')->getStyle('CN'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CO'.$cnt, '=SUBTOTAL(9,CO8:CO'.($cnt-1).')')->getStyle('CO'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CP'.$cnt, '=SUBTOTAL(9,CP8:CP'.($cnt-1).')')->getStyle('CP'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CQ'.$cnt, '=SUBTOTAL(9,CQ8:CQ'.($cnt-1).')')->getStyle('CQ'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CR'.$cnt, '=SUBTOTAL(9,CR8:CR'.($cnt-1).')')->getStyle('CR'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CS'.$cnt, '=SUBTOTAL(9,CS8:CS'.($cnt-1).')')->getStyle('CS'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CT'.$cnt, '=SUBTOTAL(9,CT8:CT'.($cnt-1).')')->getStyle('CT'.$cnt)->applyFromArray($styleArray01);	
$objPHPExcel->getActiveSheet()->setCellValue('CU'.$cnt, '=SUBTOTAL(9,CU8:CU'.($cnt-1).')')->getStyle('CU'.$cnt)->applyFromArray($styleArray01);	

}

$sql_n_mqt = "SELECT `dr`.relase AS `drrelease`, REPLACE(`dpr`.prname, '&nbsp;', '') AS `dpproduct`,`pf`.prfamname AS `prfam`,`t`.name AS `techno`,
eol.text AS elotxt,	ntag.tag AS tag, ntag.descr AS ntagdescr, `dpr`.id,`cib`.`nbofnodes`,`mqtp`.`disc`,`mqtp`.`confidance`,
`mqt`.`shelf_type`, `mqt`.`shelf_type_id`, `mqt`.`shelf_type_disp`, `mqt`.`pr_com_name`, `mqt`.`c_pr_name`, `mqt`.`c_pr_model`,`mqt`. `mqt_product`, 
`mqt`.`mqt_config`, `mqt`.`s_name_id`, `mqt`.`mqt_product_inv`, `mqt`.`mqt_config_inv`, IFNULL(`mqt`.`mqt_avg`,0) as `mqt_avg`, IFNULL(`mqt`.`psp`,0) as `psp`, `mqt`.`curr`, 
`mqt`.`mngt`,`mqt`. `excep`,`mqt`.`mqt_u_name`,`cib`.`lpoints` AS `lpoints`,`dpr`.`lpoints` AS `lpointsch`,`mqt`.`warn`
FROM `itb` AS `cib`
LEFT JOIN `prrelease` AS `dr` on `cib`.`dprrelid`=`dr`.`id`
LEFT JOIN `product` AS `dpr` on `dr`.`prid`=`dpr`.`id`
LEFT JOIN `prfam` AS `pf` on `dpr`.`prfid`=`pf`.`id`
LEFT JOIN `techno` AS `t` on `pf`.`techid`=`t`.`id`
LEFT JOIN `tatstat` AS `tatst` on `cib`.`tatre`=`tatst`.`id`
LEFT JOIN `eol` AS `eol` on eol.id=`dr`.`eolstat`
LEFT JOIN `ntag` AS `ntag` on `cib`.`ntag`=`ntag`.`id`
LEFT JOIN `shelf_mqt_conv` AS mqt on `mqt`.`s_name_id`=`dpr`.id
LEFT JOIN `mqt_cust_profile` AS mqtp on (`mqtp`.`cuid`=`cib`.cuid AND `dpr`.`prname`=`mqtp`.`shtype` AND `mqtp`.`type`= '2')
WHERE `cib`.`del`='0' AND `cib`.`ntag` IN ($tags) AND (`ntag`.`mngt`='2' or `ntag`.`mngt`='3' or `mqt`.`mngt`='0' OR `mqt`.excep='1') GROUP BY `cib`.`id` ORDER BY `cib`.`ntag` ASC,`cib`.`dprrelid` ASC";
$res1 = cib_dbQuery($sql_n_mqt);
$is_man_tr=cib_dbNumRows($res1);

If ($is_man_tr) {
$tab_cnt++;
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex($tab_cnt)->setTitle("Manual Audit Details")->getTabColor()->setRGB('FF0000');
$tabsnames[]="Manual Audit Details";
//$objPHPExcel->setActiveSheetIndex(2);
$objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(85);

$objPHPExcel->getActiveSheet()->setCellValue('A1', "Care Renewal Assessment (CRA)")->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('A2', "Manual Audit Details")->getStyle('A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('A2')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->getStyle('A2')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('A3', "This table lists the Part by part breakdown found for any products that Sonar's Audit Source is Manually entered (see Column A in Summary tab)")->getStyle('A3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('A3')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('J1', "**** Changes made to this tab will NOT be reflected in the Summary tab ****")->getStyle('J1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('J1')->getFont()->setSize(12)->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->getStyle('J1')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->mergeCells('J1:L1');

$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('B7', "Item")->getStyle('B7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('C7', 'Product Name')->getStyle('C7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('D7', 'MQT Product Name')->getStyle('D7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('E7', 'MQT Product Config')->getStyle('E7')->applyFromArray($styleArray0);	
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('F7', "Warning")->getStyle('F7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('G7', 'SW Release')->getStyle('G7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('H7', 'EOL status')->getStyle('H7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('I7', 'Qty')->getStyle('I7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('J7', 'NTAG')->getStyle('J7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('K7', 'MQT Unit - ESTIMATED')->getStyle('K7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('L7', 'MQT Unit Name')->getStyle('L7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('M7', "List Price (".$runc['curr'].") EACH - ESTIMATED")->getStyle('M7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('N7', "List Price (".$runc['curr'].") TOTAL - ESTIMATED")->getStyle('N7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('O7', 'OPD%')->getStyle('O7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('P')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('P7', 'OPD% Confidence')->getStyle('P7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('Q7', "PSP (".$runc['curr'].") EACH - ESTIMATED")->getStyle('Q7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('R7', "PSP (".$runc['curr'].") TOTAL - ESTIMATED")->getStyle('R7')->applyFromArray($styleArray0);

$i=8;
$cnt=1;

while ($ticket=cib_dbFetchAssoc($res1))
{
$confi='-';
if ($ticket['confidance']==1) $confi='High';
else if ($ticket['confidance']==2) $confi='Medium';

$unitprice=$ticket['psp'];
$cur=$ticket['curr'];

if ($cur=='USD' AND $runc['curr']=='EUR') $unitprice=$unitprice*$USD_EUR;
else if ($cur=='EUR' AND $runc['curr']=='USD') $unitprice=$unitprice*$EUR_USD;	

$mqt_product_n = (!Empty ($ticket['mqt_product']) ? $ticket['mqt_product'] : 'MQT details not available');
$mqt_config_n = (!Empty ($ticket['mqt_config']) ? $ticket['mqt_config'] : 'MQT details not available');	

$nbofnodes_lp=$ticket['nbofnodes'];
if ($ticket['lpointsch']) $nbofnodes_lp=$ticket['lpoints'];

$Rel_n=explode("R",$ticket['drrelease']);
$MQT_Prod_n[]=$mqt_product_n.':'.$mqt_config_n.':'.$Rel_n[0];
$MQT_Prod_n_eol_s[$mqt_product_n.':'.$mqt_config_n.':'.$Rel_n[0]]=$ticket['elotxt'];
$MQT_Prod_n_mqt_u_name[$mqt_product_n.':'.$mqt_config_n.':'.$Rel_n[0]]=$ticket['mqt_u_name'];
//$MQT_Prod_n_mqtv[$mqt_product_n.':'.$mqt_config_n.':'.$Rel_n[0]][]=$ticket['nbofnodes']*$ticket['mqt_avg'];
$MQT_Prod_n_mqtv[$mqt_product_n.':'.$mqt_config_n.':'.$Rel_n[0]][]=$nbofnodes_lp*$ticket['mqt_avg'];
$MQT_Prod_n_psp[$mqt_product_n.':'.$mqt_config_n.':'.$Rel_n[0]][]=$nbofnodes_lp*$unitprice;
//$MQT_Prod_n_psp_d[$mqt_product_n.':'.$mqt_config_n.':'.$Rel_n[0]][]=($ticket['nbofnodes'])*($unitprice-($unitprice*($ticket['disc']/100)));
$MQT_Prod_n_psp_d[$mqt_product_n.':'.$mqt_config_n.':'.$Rel_n[0]][]=($nbofnodes_lp)*($unitprice-($unitprice*($ticket['disc']/100)));
$MQT_Prod_n_conf[$mqt_product_n.':'.$mqt_config_n.':'.$Rel_n[0]]=$confi;

$tag_n_p=$ticket['tag'];
$mqt_u_name_m=$ticket['mqt_u_name'];
$disc_m=$ticket['disc'];
$dpproduct_m=$ticket['dpproduct'];
$Node_collector_m[]=$tag_n_p.':'.$dpproduct_m.':'.$mqt_product_n.':'.$mqt_config_n.':'.$mqt_u_name_m.':'.$confi.':'.$disc_m.':'.$Rel_n[0];
$Node_collector_m_nodes[$tag_n_p.':'.$dpproduct_m.':'.$mqt_product_n.':'.$mqt_config_n.':'.$mqt_u_name_m.':'.$confi.':'.$disc_m.':'.$Rel_n[0]][]=$ticket['nbofnodes'];
$Node_collector_m_mqtv[$tag_n_p.':'.$dpproduct_m.':'.$mqt_product_n.':'.$mqt_config_n.':'.$mqt_u_name_m.':'.$confi.':'.$disc_m.':'.$Rel_n[0]][]=$nbofnodes_lp*$ticket['mqt_avg'];
$Node_collector_m_psp[$tag_n_p.':'.$dpproduct_m.':'.$mqt_product_n.':'.$mqt_config_n.':'.$mqt_u_name_m.':'.$confi.':'.$disc_m.':'.$Rel_n[0]][]=$nbofnodes_lp*$unitprice;


$objPHPExcel->getActiveSheet()->setCellValue('B'.$i, $cnt)->getStyle('B'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('C'.$i, $ticket['dpproduct'])->getStyle('C'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('D'.$i, $mqt_product_n)->getStyle('D'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('E'.$i, $mqt_config_n)->getStyle('E'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('F'.$i, $ticket['warn'])->getStyle('F'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
//$objPHPExcel->getActiveSheet()->setCellValue('G'.$i, $Rel_n[0])->getStyle('G'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('G'.$i, "'".$ticket['drrelease'])->getStyle('G'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('H'.$i, $ticket['elotxt'])->getStyle('H'.$i)->applyFromArray($styleArray01);
//$objPHPExcel->getActiveSheet()->setCellValue('I'.$i, $ticket['nbofnodes'])->getStyle('I'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('I'.$i, $ticket['nbofnodes'])->getStyle('I'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('J'.$i, $ticket['tag'])->getStyle('J'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('K'.$i, '='.$ticket['mqt_avg'].'*'.$nbofnodes_lp.'')->getStyle('K'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('K'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
$objPHPExcel->getActiveSheet()->getStyle('K'.$i)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('L'.$i, $ticket['mqt_u_name'])->getStyle('L'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('L'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);
$objPHPExcel->getActiveSheet()->setCellValue('M'.$i, $unitprice)->getStyle('M'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('M'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);
$objPHPExcel->getActiveSheet()->getStyle('M'.$i)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('N'.$i, '='.$unitprice.'*'.$nbofnodes_lp.'')->getStyle('N'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('N'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);
$objPHPExcel->getActiveSheet()->getStyle('N'.$i)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('O'.$i, '='.$ticket['disc']/100)->getStyle('O'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('O'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
$objPHPExcel->getActiveSheet()->setCellValue('P'.$i, $confi)->getStyle('P'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('Q'.$i, '=M'.$i.'-(M'.$i.'*O'.$i.')')->getStyle('Q'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('Q'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);
$objPHPExcel->getActiveSheet()->getStyle('Q'.$i)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('R'.$i, '=Q'.$i.'*'.$nbofnodes_lp.'')->getStyle('R'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('R'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED3);
$objPHPExcel->getActiveSheet()->getStyle('R'.$i)->getFont()->getColor()->setRGB('0070C0');

$i++;
$cnt++;
}
}

$objPHPExcel->setActiveSheetIndex(1);
$objPHPExcel->getActiveSheet()->setCellValue('A1', "Care Renewal Assessment (CRA)")->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('R1', "In Warranty (IW) and Out Of Warranty (OOW) calculated from manufacturing date")->getStyle('R1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('A2', "Summary")->getStyle('A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('A2')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->getStyle('A2')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->setCellValue('AV1', "If you are using the \"Special Case\" setting (Method 4) in the Renewals Path,  the dataset below will be imported.  No Previous Contract data will be imported.")->getStyle('AV1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('AV1')->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->getStyle('AV1')->getFont()->setBold( true );

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('A7', "Sonar Audit Source")->getStyle('A7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('B7', "Item")->getStyle('B7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('C7', 'MQT Product Name')->getStyle('C7')->applyFromArray($styleArray0);
//$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth('20');
$objPHPExcel->getActiveSheet()->setCellValue('D7', 'MQT Product Config')->getStyle('D7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('E7', 'MQT Unit Name')->getStyle('E7')->applyFromArray($styleArray0);	
//$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('F7', 'SW Release at Contract Start')->getStyle('F7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('G7', 'SSL Stage at Contract Start')->getStyle('G7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('H7', 'SSL Stage at Contract End')->getStyle('H7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('I7', '')->getStyle('I7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('J7', '')->getStyle('J7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('K7', '')->getStyle('K7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('L7', 'MQT Unit Qty')->getStyle('L7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('M7', "PSP (".$runc['curr'].")")->getStyle('M7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth('14');
//$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('N7', "GLP (".$runc['curr'].")")->getStyle('N7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('O7', 'PSP Confidence')->getStyle('O7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('P7', 'OPD%')->getStyle('P7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('P')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('Q7', 'OPD% Confidence')->getStyle('Q7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('R7', 'MQT Unit Qty')->getStyle('R7')->applyFromArray($styleArray0);
//$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('S7', 'MQT Unit Qty')->getStyle('S7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('S')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('T7', 'MQT Unit Qty')->getStyle('T7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('T')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('U7', "PSP (".$runc['curr'].")")->getStyle('U7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('U')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('V7', 'LP')->getStyle('V7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('V')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('W7', 'MQT Unit Qty')->getStyle('W7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('W')->setAutoSize(true);
//$objPHPExcel->getActiveSheet()->mergeCells('R6:S6');
$objPHPExcel->getActiveSheet()->setCellValue('R6', "TS")->getStyle('R6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('S6', "RES")->getStyle('S6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->mergeCells('L6:Q6');
$objPHPExcel->getActiveSheet()->setCellValue('L6', "")->getStyle('L6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('M6', "")->getStyle('M6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('N6', "")->getStyle('N6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('O6', "")->getStyle('O6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('P6', "")->getStyle('P6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('Q6', "")->getStyle('Q6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('L4', "Previous Contract, started $mtcpst")->getStyle('L4:X4')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getStyle('L4')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->mergeCells('L4:X4');
$objPHPExcel->getActiveSheet()->getStyle('L4:X4')->getFill()->getStartColor()->setRGB('8DB4E2');
$objPHPExcel->getActiveSheet()->mergeCells('L5:Q5');
$objPHPExcel->getActiveSheet()->setCellValue('L5', "Prior IB Total")->getStyle('L5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('M5', "")->getStyle('M5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('N5', "")->getStyle('N5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('O5', "")->getStyle('O5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('P5', "")->getStyle('P5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('Q5', "")->getStyle('Q5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->mergeCells('R5:S5');
$objPHPExcel->getActiveSheet()->setCellValue('R5', "Prior IB (IW)")->getStyle('R5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('S5', "")->getStyle('S5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('X7', 'PSP')->getStyle('X7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('Y7', 'LP')->getStyle('Y7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setVisible(false);
$objPHPExcel->getActiveSheet()->mergeCells('T6:U6');
$objPHPExcel->getActiveSheet()->setCellValue('T6', "TS")->getStyle('T6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('U6', "")->getStyle('U6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->mergeCells('W6:X6');
$objPHPExcel->getActiveSheet()->setCellValue('W6', "RES")->getStyle('W6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('X6', "")->getStyle('X6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->mergeCells('T5:X5');
$objPHPExcel->getActiveSheet()->setCellValue('T5', "Prior IB (OOW)")->getStyle('T5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('U5', "")->getStyle('U5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('W5', "")->getStyle('W5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('X5', "")->getStyle('X5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('Z7', "MQT Unit Qty")->getStyle('Z7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->mergeCells('Z4:AD4');
$objPHPExcel->getActiveSheet()->setCellValue('Z4', "Increment after $mtcpst.")->getStyle('Z4:AD4')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getStyle('Z4')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->getStyle('Z4:AD4')->getFill()->getStartColor()->setRGB('8DB4E2');
$objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AA7', "PSP (".$runc['curr'].")")->getStyle('AA7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AB7', "GLP (".$runc['curr'].")")->getStyle('AB7')->applyFromArray($styleArray0);
//$objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AC')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('AC')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('AC7', "PSP Confidence")->getStyle('AC7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('AD')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('AD7', "OPD%")->getStyle('AD7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('AE')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('AE7', "OPD% Confidence")->getStyle('AE7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->mergeCells('Z5:AE6');
$objPHPExcel->getActiveSheet()->setCellValue('Z5', "Change in IB")->getStyle('Z5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AA5', "")->getStyle('AA5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AB5', "")->getStyle('AB5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AC5', "")->getStyle('AC5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AD5', "")->getStyle('AD5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('AE')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('AE5', "")->getStyle('AE5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('Z6', "")->getStyle('Z6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AA6', "")->getStyle('AA6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AB6', "")->getStyle('AB6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AC6', "")->getStyle('AC6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AD6', "")->getStyle('AD6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AE6', "")->getStyle('AE6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AF4', "New Contract, starting $mtcstd. (All data below represents the anticipated IB state on this date.)")->getStyle('AF4:AR4')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getStyle('AF4')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->mergeCells('AF4:AR4');
$objPHPExcel->getActiveSheet()->getStyle('AF4:AR4')->getFill()->getStartColor()->setRGB('8DB4E2');
$objPHPExcel->getActiveSheet()->getColumnDimension('AF')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('AF7', "MQT Unit Qty")->getStyle('AF7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('AG')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AG7', "PSP (".$runc['curr'].")")->getStyle('AG7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('AH')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AH7', "GLP (".$runc['curr'].")")->getStyle('AH7')->applyFromArray($styleArray0);
//$objPHPExcel->getActiveSheet()->getColumnDimension('AH')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AI')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('AI')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('AI7', "PSP Confidence")->getStyle('AI7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('AJ')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('AJ7', "OPD%")->getStyle('AJ7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('AK')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('AK7', "OPD% Confidence")->getStyle('AK7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('AL')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('AL7', "MQT Unit Qty")->getStyle('AL7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('AK')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('AM7', "MQT Unit Qty")->getStyle('AM7')->applyFromArray($styleArray0);
//$objPHPExcel->getActiveSheet()->mergeCells('AL6:AM6');
$objPHPExcel->getActiveSheet()->setCellValue('AL6', "TS")->getStyle('AL6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AM6', "RES")->getStyle('AM6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->mergeCells('AF5:AK5');
$objPHPExcel->getActiveSheet()->setCellValue('AF5', "New Contract IB Total")->getStyle('AF5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AG5', "")->getStyle('AG5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AH5', "")->getStyle('AH5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AI5', "")->getStyle('AI5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AJ5', "")->getStyle('AJ5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AK5', "")->getStyle('AK5')->applyFromArray($styleArray0);
//$objPHPExcel->getActiveSheet()->setCellValue('AG5', "")->getStyle('AG5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->mergeCells('AL5:AM5');
$objPHPExcel->getActiveSheet()->setCellValue('AL5', "New Contract IB (IW)")->getStyle('AL5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AM5', "")->getStyle('AM5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->mergeCells('AF6:AK6');
$objPHPExcel->getActiveSheet()->setCellValue('AF6', "")->getStyle('AF6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AG6', "")->getStyle('AG6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AH6', "")->getStyle('AH6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AI6', "")->getStyle('AI6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AJ6', "")->getStyle('AJ6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AK6', "")->getStyle('AK6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('AN')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('AN7', "MQT Unit Qty")->getStyle('AN7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('AO')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AO7', "PSP (".$runc['curr'].")")->getStyle('AO7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('AP')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AQ')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->setCellValue('AQ7', "MQT Unit Qty")->getStyle('AQ7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('AR')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AR7', "PSP (".$runc['curr'].")")->getStyle('AR7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('AS')->setVisible(false);
$objPHPExcel->getActiveSheet()->mergeCells('AN6:AO6');
$objPHPExcel->getActiveSheet()->setCellValue('AN6', "TS")->getStyle('AN6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AO6', "")->getStyle('AO6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->mergeCells('AQ6:AR6');
$objPHPExcel->getActiveSheet()->setCellValue('AQ6', "RES")->getStyle('AQ6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AR6', "")->getStyle('AR6')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->mergeCells('AN5:AR5');
$objPHPExcel->getActiveSheet()->setCellValue('AN5', "New Contract IB (OOW)")->getStyle('AN5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AO5', "")->getStyle('AO5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AQ5', "")->getStyle('AQ5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->setCellValue('AR5', "")->getStyle('AR5')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->mergeCells('AV5:AZ6');
$objPHPExcel->getActiveSheet()->setCellValue('AV5', "Total IB.  (This data will be imported into \"Change in IB\" section of GUI)")->getStyle('AV5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AW5', "")->getStyle('AW5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AV6', "")->getStyle('AV6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AZ6', "")->getStyle('AZ6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AX5', "")->getStyle('AX5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AY5', "")->getStyle('AY5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AZ5', "")->getStyle('AZ5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('BA5:BG5');
$objPHPExcel->getActiveSheet()->setCellValue('BA5', "Current IB Total")->getStyle('BA5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BB5', "")->getStyle('BB5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BC5', "")->getStyle('BC5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BD5', "")->getStyle('BD5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BE5', "")->getStyle('BE5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BF5', "")->getStyle('BF5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BG5', "")->getStyle('BG5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('BA6:BE6');
$objPHPExcel->getActiveSheet()->setCellValue('BA6', "")->getStyle('BA6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BB6', "")->getStyle('BB6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BC6', "")->getStyle('BC6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BD6', "")->getStyle('BD6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BE6', "")->getStyle('BE6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('BF6:BG6');
$objPHPExcel->getActiveSheet()->setCellValue('BF6', "Qty IW")->getStyle('BF6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BG6', "")->getStyle('BG6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('BH6:BI6');
$objPHPExcel->getActiveSheet()->setCellValue('BH6', "TS")->getStyle('BH6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BI6', "")->getStyle('BI6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('BH5:BK5');
$objPHPExcel->getActiveSheet()->setCellValue('BH5', "New Contract IB (OOW)")->getStyle('BH5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BI5', "")->getStyle('BI5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BJ5', "")->getStyle('BJ5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BK5', "")->getStyle('BK5')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->mergeCells('BJ6:BK6');
$objPHPExcel->getActiveSheet()->setCellValue('BJ6', "RES")->getStyle('BJ6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BK6', "")->getStyle('BK6')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BA4', "New Contract, starting $mtcstd. (All data below represents the anticipated IB state on this date.)")->getStyle('BA4:BK4')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getStyle('BA4')->getFont()->setBold( true );
$objPHPExcel->getActiveSheet()->mergeCells('BA4:BK4');
$objPHPExcel->getActiveSheet()->getStyle('BA4:BK4')->getFill()->getStartColor()->setRGB('FABF8F');

$objPHPExcel->getActiveSheet()->getColumnDimension('AV')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AW')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AV')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AX')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AY')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('AZ')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BA')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BB')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BC')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BD')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BE')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BF')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BG')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BH')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BI')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BJ')->setVisible(false);
$objPHPExcel->getActiveSheet()->getColumnDimension('BK')->setVisible(false);

$objPHPExcel->getActiveSheet()->setCellValue('AV7', "MQT Unit Qty")->getStyle('AV7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AW7', "PSP (".$runc['curr'].")")->getStyle('AW7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('AW')->setWidth('14');
$objPHPExcel->getActiveSheet()->getColumnDimension('AV')->setWidth('14');
$objPHPExcel->getActiveSheet()->getColumnDimension('AX')->setWidth('14');
$objPHPExcel->getActiveSheet()->getColumnDimension('AY')->setWidth('14');
$objPHPExcel->getActiveSheet()->getColumnDimension('AZ')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AX7', "PSP Confidence")->getStyle('AX7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AY7', "OPD%")->getStyle('AY7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('AZ7', "OPD% Confidence")->getStyle('AZ7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BA7', "MQT Unit Qty")->getStyle('BA7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BB7', "PSP (".$runc['curr'].")")->getStyle('BB7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BB')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('BC7', "PSP Confidence")->getStyle('BC7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BD7', "OPD%")->getStyle('BD7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BE7', "OPD% Confidence")->getStyle('BE7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BF7', "TS")->getStyle('BF7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BG7', "RES")->getStyle('BG7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BH7', "MQT Unit Qty")->getStyle('BH7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BI7', "PSP (".$runc['curr'].")")->getStyle('BI7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BJ7', "MQT Unit Qty")->getStyle('BJ7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->setCellValue('BK7', "PSP (".$runc['curr'].")")->getStyle('BK7')->applyFromArray($styleArray02);
$objPHPExcel->getActiveSheet()->getColumnDimension('BI')->setWidth('14');
$objPHPExcel->getActiveSheet()->getColumnDimension('BK')->setWidth('14');
$objPHPExcel->getActiveSheet()->setCellValue('AT7', "Validation Errors")->getStyle('AT7')->applyFromArray($styleArray0);
$objPHPExcel->getActiveSheet()->getColumnDimension('AT')->setWidth('14');

$i=8;
$cnt=1;

//print_r($MQT_Qty);
//print_r($MQT_Uname);
//echo '</br>';

If ($is_el_tr) {
$M=0;
$U=0;
$AA=0;
$AG=0;

foreach (array_unique($MQT_Prod) as $tmp)
{
//$OPD=array_sum($MQT_OPD[$tmp])/sizeof($MQT_OPD[$tmp]);
if (array_sum($MQT_AA_FP[$tmp])>0)
$OPD=(array_sum($MQT_AA_FP[$tmp])-array_sum($MQT_AA[$tmp]))/array_sum($MQT_AA_FP[$tmp]);
else $OPD=0;

if (array_sum($MQT_AK[$tmp])>0)
$OPD_AK=(array_sum($MQT_AK_FP[$tmp])-array_sum($MQT_AK[$tmp]))/array_sum($MQT_AK_FP[$tmp]);
else $OPD_AK=0;

if ((array_sum($MQT_AK[$tmp])-array_sum($MQT_AA[$tmp]))>0)
$OPD_AD=(((array_sum($MQT_AK_FP[$tmp])-array_sum($MQT_AA_FP[$tmp]))-(array_sum($MQT_AK[$tmp])-array_sum($MQT_AA[$tmp])))/(array_sum($MQT_AK_FP[$tmp])-array_sum($MQT_AA_FP[$tmp])));
else $OPD_AD=0;

$MQT_row=explode(":",$tmp);
//$objPHPExcel->setActiveSheetIndex(1);
$objPHPExcel->getActiveSheet()->setCellValue('A'.$i, 'Electronic')->getStyle('A'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('B'.$i, $cnt)->getStyle('B'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('C'.$i, $MQT_row[0])->getStyle('C'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('D'.$i, $MQT_row[1])->getStyle('D'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('D'.$i)->getAlignment()->setWrapText(false);
$objPHPExcel->getActiveSheet()->setCellValue('E'.$i, $MQT_Uname[$tmp])->getStyle('E'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('F'.$i, '\''.$MQT_row[2])->getStyle('F'.$i)->applyFromArray($styleArray01G);
$objPHPExcel->getActiveSheet()->setCellValue('G'.$i, '')->getStyle('G'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('H'.$i, '')->getStyle('H'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('I'.$i, '')->getStyle('I'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('J'.$i, '')->getStyle('J'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('K'.$i, '')->getStyle('K'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('L'.$i, round(array_sum($MQT_AD[$tmp])))->getStyle('L'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objValidation = $objPHPExcel->getActiveSheet()->getCell('L'.$i)->getDataValidation();
$objValidation->setType( PHPExcel_Cell_DataValidation::TYPE_WHOLE );
$objValidation->setErrorStyle( PHPExcel_Cell_DataValidation::STYLE_STOP );
$objValidation->setOperator( PHPExcel_Cell_DataValidation::OPERATOR_GREATERTHANOREQUAL);
$objValidation->setAllowBlank(true);
$objValidation->setShowInputMessage(true);
$objValidation->setShowErrorMessage(true);
$objValidation->setErrorTitle('Input error');
$objValidation->setError('Only digits');
//$objValidation->setPromptTitle('Allowed input');
//$objValidation->setPrompt('Only 1 and 999999999 are allowed.');
$objValidation->setFormula1(0);
//$objValidation->setFormula2(999999999);
$objPHPExcel->getActiveSheet()->getStyle('L'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	

$M=array_sum($MQT_AA_FP[$tmp])-(array_sum($MQT_AA_FP[$tmp])*$OPD);

//$objPHPExcel->getActiveSheet()->setCellValue('M'.$i, '=N'.$i.'-(N'.$i.'*P'.$i.')')->getStyle('M'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->setCellValue('M'.$i, $M)->getStyle('M'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$objValidation = $objPHPExcel->getActiveSheet()->getCell('M'.$i)->getDataValidation();
$objValidation->setType( PHPExcel_Cell_DataValidation::TYPE_DECIMAL );
$objValidation->setErrorStyle( PHPExcel_Cell_DataValidation::STYLE_STOP );
$objValidation->setOperator( PHPExcel_Cell_DataValidation::OPERATOR_GREATERTHANOREQUAL);
$objValidation->setAllowBlank(true);
$objValidation->setShowInputMessage(true);
$objValidation->setShowErrorMessage(true);
$objValidation->setErrorTitle('Input error');
$objValidation->setError('Only digits');
//$objValidation->setPromptTitle('Allowed input');
//$objValidation->setPrompt('Only 1 and 999999999 are allowed.');
$objValidation->setFormula1(0);
//$objValidation->setFormula2(999999999);
$objPHPExcel->getActiveSheet()->getStyle('M'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	

$objPHPExcel->getActiveSheet()->setCellValue('N'.$i, array_sum($MQT_AA_FP[$tmp]))->getStyle('N'.$i)->applyFromArray($styleArray01G);
$objPHPExcel->getActiveSheet()->getStyle('N'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
$objPHPExcel->getActiveSheet()->setCellValue('O'.$i, $MQT_O[$tmp])->getStyle('O'.$i)->applyFromArray($styleArray01G);
$objValidation = $objPHPExcel->getActiveSheet()->getCell('O'.$i)->getDataValidation();
$objValidation->setType( PHPExcel_Cell_DataValidation::TYPE_LIST );
$objValidation->setErrorStyle( PHPExcel_Cell_DataValidation::STYLE_STOP );
$objValidation->setAllowBlank(true);
$objValidation->setShowInputMessage(true);
$objValidation->setShowErrorMessage(true);
$objValidation->setShowDropDown(true);
$objValidation->setErrorTitle('Input error');
$objValidation->setError('Value is not in list.');
//$validation->setPromptTitle('Pick from list');
//$validation->setPrompt('Please pick a value from the drop-down list.');
$objValidation->setFormula1('"High,Medium"');
$objValidation = $objPHPExcel->getActiveSheet()->getCell('O'.$i)->getDataValidation();
$objPHPExcel->getActiveSheet()->setCellValue('P'.$i, $OPD)->getStyle('P'.$i)->applyFromArray($styleArray01G);
$objValidation = $objPHPExcel->getActiveSheet()->getCell('P'.$i)->getDataValidation();
$objValidation->setType( PHPExcel_Cell_DataValidation::TYPE_DECIMAL );
$objValidation->setErrorStyle( PHPExcel_Cell_DataValidation::STYLE_STOP );
$objValidation->setAllowBlank(true);
$objValidation->setShowInputMessage(true);
$objValidation->setShowErrorMessage(true);
$objValidation->setErrorTitle('Input error');
$objValidation->setError('Only decimal values');
//$objValidation->setPromptTitle('Allowed input');
//$objValidation->setPrompt('Only 1 and 999999999 are allowed.');
$objValidation->setFormula1(0);
$objValidation->setFormula2(1);
$objPHPExcel->getActiveSheet()->getStyle('P'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);

If ($OPD==0.00){
$objPHPExcel->getActiveSheet()->getStyle('M'.$i)->getFont()->getColor()->setRGB('FF0000');
$objPHPExcel->getActiveSheet()->getStyle('P'.$i)->getFont()->getColor()->setRGB('FF0000');
}

$objPHPExcel->getActiveSheet()->setCellValue('Q'.$i, $MQT_O[$tmp])->getStyle('Q'.$i)->applyFromArray($styleArray01G);
$objPHPExcel->getActiveSheet()->setCellValue('R'.$i, round(array_sum($MQT_AB[$tmp])))->getStyle('R'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('R'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
$objPHPExcel->getActiveSheet()->setCellValue('S'.$i, round(array_sum($MQT_BP[$tmp])))->getStyle('S'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('S'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
$objPHPExcel->getActiveSheet()->setCellValue('T'.$i, round(array_sum($MQT_AC[$tmp])))->getStyle('T'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('T'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	

$U=array_sum($MQT_Z_FP[$tmp])-(array_sum($MQT_Z_FP[$tmp])*$OPD);
//$objPHPExcel->getActiveSheet()->setCellValue('U'.$i, '=V'.$i.'-(V'.$i.'*P'.$i.')')->getStyle('U'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->setCellValue('U'.$i, $U)->getStyle('U'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);


$objPHPExcel->getActiveSheet()->getStyle('U'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
$objPHPExcel->getActiveSheet()->setCellValue('V'.$i, array_sum($MQT_Z_FP[$tmp]))->getStyle('V'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('W'.$i, round(array_sum($MQT_BQ[$tmp])))->getStyle('W'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('W'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
$objPHPExcel->getActiveSheet()->setCellValue('X'.$i, '=Y'.$i.'-(Y'.$i.'*P'.$i.')')->getStyle('X'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('X'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
$objPHPExcel->getActiveSheet()->setCellValue('Y'.$i, array_sum($MQT_BN_FP[$tmp]))->getStyle('Y'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('Z'.$i, '=AF'.$i.'-L'.$i.'')->getStyle('Z'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('Z'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	

$AA=(array_sum($MQT_AK_FP[$tmp])-array_sum($MQT_AA_FP[$tmp]))-((array_sum($MQT_AK_FP[$tmp])-array_sum($MQT_AA_FP[$tmp]))*$OPD_AD);

//$objPHPExcel->getActiveSheet()->setCellValue('AA'.$i, '=AB'.$i.'-(AB'.$i.'*AD'.$i.')')->getStyle('AA'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->setCellValue('AA'.$i, $AA)->getStyle('AA'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$objPHPExcel->getActiveSheet()->getStyle('AA'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
$objPHPExcel->getActiveSheet()->setCellValue('AB'.$i, '=AH'.$i.'-N'.$i.'')->getStyle('AB'.$i)->applyFromArray($styleArray01G);
$objPHPExcel->getActiveSheet()->setCellValue('AC'.$i, $MQT_O[$tmp])->getStyle('AC'.$i)->applyFromArray($styleArray01G);
//$objPHPExcel->getActiveSheet()->setCellValue('AD'.$i, '=((AH'.$i.'-N'.$i.')-(AG'.$i.'-M'.$i.'))/(AH'.$i.'-N'.$i.')')->getStyle('AD'.$i)->applyFromArray($styleArray01G);
$objPHPExcel->getActiveSheet()->setCellValue('AD'.$i, $OPD_AD)->getStyle('AD'.$i)->applyFromArray($styleArray01G);
$objPHPExcel->getActiveSheet()->getStyle('AD'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
$objPHPExcel->getActiveSheet()->setCellValue('AE'.$i, $MQT_O[$tmp])->getStyle('AE'.$i)->applyFromArray($styleArray01G);
$objPHPExcel->getActiveSheet()->setCellValue('AF'.$i, round(array_sum($MQT_AN[$tmp])))->getStyle('AF'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('AF'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	

$AG=array_sum($MQT_AK_FP[$tmp])-(array_sum($MQT_AK_FP[$tmp])*$OPD_AK);
//$objPHPExcel->getActiveSheet()->setCellValue('AG'.$i, '=AH'.$i.'-(AH'.$i.'*AJ'.$i.')')->getStyle('AG'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->setCellValue('AG'.$i, $AG)->getStyle('AG'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);


$objPHPExcel->getActiveSheet()->getStyle('AG'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
$objPHPExcel->getActiveSheet()->setCellValue('AH'.$i, array_sum($MQT_AK_FP[$tmp]))->getStyle('AH'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('AH'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);
$objPHPExcel->getActiveSheet()->setCellValue('AI'.$i, $MQT_O[$tmp])->getStyle('AI'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AJ'.$i, $OPD_AK)->getStyle('AJ'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('AJ'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
$objPHPExcel->getActiveSheet()->setCellValue('AK'.$i, $MQT_O[$tmp])->getStyle('AK'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AL'.$i, round(array_sum($MQT_AL[$tmp])))->getStyle('AL'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('AL'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
$objPHPExcel->getActiveSheet()->setCellValue('AM'.$i, round(array_sum($MQT_BZ[$tmp])))->getStyle('AM'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('AM'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
$objPHPExcel->getActiveSheet()->setCellValue('AN'.$i, round(array_sum($MQT_AN2[$tmp])))->getStyle('AN'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('AN'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
$objPHPExcel->getActiveSheet()->setCellValue('AO'.$i, array_sum($MQT_AO[$tmp]))->getStyle('AO'.$i)->applyFromArray($styleArray01G);
//$objPHPExcel->getActiveSheet()->setCellValue('AO'.$i, '=AP'.$i.'-(AP'.$i.'*AJ'.$i.')')->getStyle('AO'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('AO'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
$objPHPExcel->getActiveSheet()->setCellValue('AP'.$i, array_sum($MQT_AO_FP[$tmp]))->getStyle('AP'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AQ'.$i, round(array_sum($MQT_AQ[$tmp])))->getStyle('AQ'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('AQ'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
$objPHPExcel->getActiveSheet()->setCellValue('AR'.$i, array_sum($MQT_AR[$tmp]))->getStyle('AR'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
//$objPHPExcel->getActiveSheet()->setCellValue('AR'.$i, '=AS'.$i.'-(AS'.$i.'*AJ'.$i.')')->getStyle('AR'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('AR'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
$objPHPExcel->getActiveSheet()->setCellValue('AS'.$i, array_sum($MQT_AR_FP[$tmp]))->getStyle('AS'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AV'.$i, '=AF'.$i.'')->getStyle('AV'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('AV'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->setCellValue('AX'.$i, '=AI'.$i.'')->getStyle('AX'.$i)->applyFromArray($styleArray01G);
$objPHPExcel->getActiveSheet()->setCellValue('AW'.$i, '=AG'.$i.'')->getStyle('AW'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('AW'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
$objPHPExcel->getActiveSheet()->setCellValue('AY'.$i, '=AJ'.$i.'')->getStyle('AY'.$i)->applyFromArray($styleArray01G);
$objPHPExcel->getActiveSheet()->getStyle('AY'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
$objPHPExcel->getActiveSheet()->setCellValue('AZ'.$i, '=AK'.$i.'')->getStyle('AZ'.$i)->applyFromArray($styleArray01G);
$objPHPExcel->getActiveSheet()->setCellValue('BA'.$i, '=AF'.$i.'')->getStyle('BA'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('BA'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->setCellValue('BB'.$i, '=AG'.$i.'')->getStyle('BB'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('BB'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);
$objPHPExcel->getActiveSheet()->setCellValue('BC'.$i, '=AI'.$i.'')->getStyle('BC'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('BD'.$i, '=AJ'.$i.'')->getStyle('BD'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('BD'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
$objPHPExcel->getActiveSheet()->setCellValue('BE'.$i, '=AK'.$i.'')->getStyle('BE'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('BF'.$i, '=AL'.$i.'')->getStyle('BF'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('BF'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
$objPHPExcel->getActiveSheet()->setCellValue('BG'.$i, '=AM'.$i.'')->getStyle('BG'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('BG'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->setCellValue('BH'.$i, '=AN'.$i.'')->getStyle('BH'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('BH'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->setCellValue('BI'.$i, '=AO'.$i.'')->getStyle('BI'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('BI'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);
$objPHPExcel->getActiveSheet()->setCellValue('BJ'.$i, '=AQ'.$i.'')->getStyle('BJ'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('BJ'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->setCellValue('BK'.$i, '=AR'.$i.'')->getStyle('BK'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('BK'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);
//$objPHPExcel->getActiveSheet()->setCellValue('BU'.$i, '=IF(SUM(L'.$i.'+Z'.$i.')<>AF'.$i.',"E1, ","")')->getStyle('BU'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('BU'.$i, '=IF(OR(SUM(L'.$i.'+Z'.$i.')>AF'.$i.',SUM(L'.$i.'+Z'.$i.')<AF'.$i.'),"E1, ","")')->getStyle('BU'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getColumnDimension('BU')->setVisible(false);
//$objPHPExcel->getActiveSheet()->setCellValue('BV'.$i, '=IF(SUM(M'.$i.'+AA'.$i.')<>AG'.$i.',"E2, ","")')->getStyle('BV'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('BV'.$i, '=IF(OR(SUM(M'.$i.'+AA'.$i.')>AG'.$i.'+1,SUM(M'.$i.'+AA'.$i.')<AG'.$i.'-1),"E2, ","")')->getStyle('BV'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getColumnDimension('BV')->setVisible(false);
$objPHPExcel->getActiveSheet()->setCellValue('AT'.$i, '=CONCATENATE(BU'.$i.',BV'.$i.')')->getStyle('AT'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('AT'.$i)->getFont()->setBold(true)->getColor()->setRGB('FF0000');

//echo $tmp.' ';
//print_r(array_sum($MQT_Qty[$tmp]));echo ' '; print_r(array_sum($MQT_PSP[$tmp]));echo ' '.$OPD;echo '</br>';
$i++;
$cnt++;
}
$cnt_s=$i;

//$cnt_ns=1;
//$i_ns=8;
$objPHPExcel->setActiveSheetIndex(2);
foreach ($Node_collector as $tmp)
{
$Node_s_row=explode(":",$tmp);
$objPHPExcel->getActiveSheet()->setCellValue('B'.$i_ns, $cnt_ns)->getStyle('B'.$i_ns)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('C'.$i_ns, 'Electronic')->getStyle('C'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('D'.$i_ns, $Node_s_row[0])->getStyle('D'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('E'.$i_ns, $Node_s_row[11])->getStyle('E'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('F'.$i_ns, $Node_s_row[2])->getStyle('F'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('G'.$i_ns, '')->getStyle('G'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('H'.$i_ns, $Node_s_row[3])->getStyle('H'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('I'.$i_ns, $Node_s_row[4])->getStyle('I'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('J'.$i_ns, '\''.$Node_s_row[7])->getStyle('J'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('J'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->setCellValue('K'.$i_ns, '')->getStyle('K'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('L'.$i_ns, $Node_s_row[8])->getStyle('L'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
//$objPHPExcel->getActiveSheet()->setCellValue('M'.$i_ns, '')->getStyle('M'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('N'.$i_ns, $Node_s_row[10])->getStyle('N'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('N'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->setCellValue('O'.$i_ns, $Node_s_row[9])->getStyle('O'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('O'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->setCellValue('P'.$i_ns, $Node_s_row[12])->getStyle('P'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('Q'.$i_ns, '')->getStyle('Q'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('R'.$i_ns, $Node_s_row[13])->getStyle('R'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('R'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
$objPHPExcel->getActiveSheet()->setCellValue('S'.$i_ns, $Node_s_row[6]/100)->getStyle('S'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('S'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
$objPHPExcel->getActiveSheet()->setCellValue('T'.$i_ns, (1-($Node_s_row[6]/100))*$Node_s_row[13])->getStyle('T'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('T'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
$objPHPExcel->getActiveSheet()->setCellValue('U'.$i_ns, $Node_s_row[5])->getStyle('U'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('M'.$i_ns, $Node_s_row[14])->getStyle('M'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

$cnt_ns++;
$i_ns++;
}
}

If ($is_man_tr) {

foreach (array_unique($MQT_Prod_n) as $tmp)
{
if (array_sum($MQT_Prod_n_psp[$tmp])>0)
$OPD=(array_sum($MQT_Prod_n_psp[$tmp])-array_sum($MQT_Prod_n_psp_d[$tmp]))/array_sum($MQT_Prod_n_psp[$tmp]);
else $OPD=0;
$MQT_n_row=explode(":",$tmp);
$objPHPExcel->setActiveSheetIndex(1);
$objPHPExcel->getActiveSheet()->setCellValue('A'.$i, 'Manual')->getStyle('A'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('B'.$i, $cnt)->getStyle('B'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('C'.$i, $MQT_n_row[0])->getStyle('C'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('D'.$i, $MQT_n_row[1])->getStyle('D'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('D'.$i)->getAlignment()->setWrapText(false);
$objPHPExcel->getActiveSheet()->setCellValue('E'.$i, $MQT_Prod_n_mqt_u_name[$tmp])->getStyle('E'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('G'.$i, $MQT_Prod_n_eol_s[$tmp])->getStyle('G'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('H'.$i, '')->getStyle('H'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('F'.$i, '\''.$MQT_n_row[2])->getStyle('F'.$i)->applyFromArray($styleArray01G);
$objPHPExcel->getActiveSheet()->setCellValue('L'.$i, '')->getStyle('L'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('L'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('M'.$i, '')->getStyle('M'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('M'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('N'.$i, '')->getStyle('N'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('N'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('O'.$i, '')->getStyle('O'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('O'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('P'.$i, '')->getStyle('P'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('P'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('Q'.$i, '')->getStyle('Q'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('Q'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('R'.$i, '')->getStyle('R'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('R'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('S'.$i, '')->getStyle('S'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('S'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('T'.$i, '')->getStyle('T'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('T'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('U'.$i, '')->getStyle('U'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('U'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('W'.$i, '')->getStyle('W'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('W'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('X'.$i, '')->getStyle('X'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('X'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('Y'.$i, '')->getStyle('Y'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('Y'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('Z'.$i, '')->getStyle('Z'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('Z'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('AA'.$i, '')->getStyle('AA'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('AA'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('AB'.$i, '')->getStyle('AB'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('AB'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('AC'.$i, '')->getStyle('AC'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('AC'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('AD'.$i, '')->getStyle('AD'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('AD'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('AE'.$i, '')->getStyle('AE'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('AE'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('AL'.$i, '')->getStyle('AL'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('AL'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('AM'.$i, '')->getStyle('AM'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('AM'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('BF'.$i, '')->getStyle('BF'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('BF'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('BG'.$i, '')->getStyle('BG'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('BG'.$i)->getFill()->getStartColor()->setRGB('BFBFBF');
$objPHPExcel->getActiveSheet()->setCellValue('AF'.$i, array_sum($MQT_Prod_n_mqtv[$tmp]))->getStyle('AF'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('AF'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->getStyle('AF'.$i)->getFont()->getColor()->setRGB('0070C0');

$AG=array_sum($MQT_Prod_n_psp[$tmp])-(array_sum($MQT_Prod_n_psp[$tmp])*$OPD);
//$objPHPExcel->getActiveSheet()->setCellValue('AG'.$i, '=AH'.$i.'-(AH'.$i.'*AJ'.$i.')')->getStyle('AG'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->setCellValue('AG'.$i, $AG)->getStyle('AG'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$objPHPExcel->getActiveSheet()->getStyle('AG'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);
$objPHPExcel->getActiveSheet()->getStyle('AG'.$i)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('AH'.$i, array_sum($MQT_Prod_n_psp[$tmp]))->getStyle('AH'.$i)->applyFromArray($styleArray01);

$objPHPExcel->getActiveSheet()->getStyle('AH'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);
$objPHPExcel->getActiveSheet()->getStyle('AH'.$i)->getFont()->getColor()->setRGB('0070C0');

$objPHPExcel->getActiveSheet()->setCellValue('AI'.$i, $MQT_Prod_n_conf[$tmp])->getStyle('AI'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AJ'.$i, $OPD)->getStyle('AJ'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('AK'.$i, $MQT_Prod_n_conf[$tmp])->getStyle('AK'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('AJ'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
$objPHPExcel->getActiveSheet()->setCellValue('AN'.$i, '=AF'.$i.'')->getStyle('AN'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('AN'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->getStyle('AN'.$i)->getFont()->getColor()->setRGB('0070C0');

//$objPHPExcel->getActiveSheet()->setCellValue('AO'.$i, '=AG'.$i.'')->getStyle('AO'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->setCellValue('AO'.$i, $AG)->getStyle('AO'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

$objPHPExcel->getActiveSheet()->getStyle('AO'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);
$objPHPExcel->getActiveSheet()->getStyle('AO'.$i)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('AQ'.$i, '=AF'.$i.'')->getStyle('AQ'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('AQ'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->getStyle('AQ'.$i)->getFont()->getColor()->setRGB('0070C0');

//$objPHPExcel->getActiveSheet()->setCellValue('AR'.$i, '=AG'.$i.'')->getStyle('AR'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->setCellValue('AR'.$i, $AG)->getStyle('AR'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);


$objPHPExcel->getActiveSheet()->getStyle('AR'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);
$objPHPExcel->getActiveSheet()->getStyle('AR'.$i)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('AV'.$i, '=AF'.$i.'')->getStyle('AV'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('AV'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->getStyle('AV'.$i)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('AW'.$i, '=AG'.$i.'')->getStyle('AW'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('AW'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);
$objPHPExcel->getActiveSheet()->getStyle('AW'.$i)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('AX'.$i, '=AI'.$i.'')->getStyle('AX'.$i)->applyFromArray($styleArray01G);
$objPHPExcel->getActiveSheet()->setCellValue('AY'.$i, '=AJ'.$i.'')->getStyle('AY'.$i)->applyFromArray($styleArray01G);
$objPHPExcel->getActiveSheet()->getStyle('AY'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
$objPHPExcel->getActiveSheet()->setCellValue('AZ'.$i, '=AK'.$i.'')->getStyle('AZ'.$i)->applyFromArray($styleArray01G);
$objPHPExcel->getActiveSheet()->setCellValue('BA'.$i, '=AF'.$i.'')->getStyle('BA'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('BA'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->getStyle('BA'.$i)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('BB'.$i, '=AG'.$i.'')->getStyle('BB'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('BB'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);
$objPHPExcel->getActiveSheet()->getStyle('BB'.$i)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('BC'.$i, '=AI'.$i.'')->getStyle('BC'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('BD'.$i, '=AJ'.$i.'')->getStyle('BD'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('BD'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
$objPHPExcel->getActiveSheet()->setCellValue('BE'.$i, '=AK'.$i.'')->getStyle('BE'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('BH'.$i, '=AN'.$i.'')->getStyle('BH'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('BH'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->getStyle('BH'.$i)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('BI'.$i, '=AO'.$i.'')->getStyle('BI'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('BI'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);
$objPHPExcel->getActiveSheet()->getStyle('BI'.$i)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('BJ'.$i, '=AQ'.$i.'')->getStyle('BJ'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('BJ'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->setCellValue('BK'.$i, '=AR'.$i.'')->getStyle('BK'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('BK'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);
$objPHPExcel->getActiveSheet()->getStyle('BK'.$i)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('AT'.$i, '=CONCATENATE(BU'.$i.',BV'.$i.')')->getStyle('AT'.$i)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->getStyle('AT'.$i)->getFont()->setBold(true)->getColor()->setRGB('FF0000');
$i++;
$cnt++;
}
}
	if ($is_alsm_ok)
	{
		$objPHPExcel->setActiveSheetIndex(1);
		foreach (array_unique($MQT_ALSM_SUM) as $tmp)
		{
			$mqt_alsm_row=explode(":",$tmp);
			$objPHPExcel->getActiveSheet()->setCellValue('A'.$i, 'ALSM')->getStyle('A'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->setCellValue('B'.$i, $cnt)->getStyle('B'.$i)->applyFromArray($styleArray01);
			$objPHPExcel->getActiveSheet()->setCellValue('C'.$i, $mqt_alsm_row[0])->getStyle('C'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->setCellValue('D'.$i, $mqt_alsm_row[1])->getStyle('D'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			//$objPHPExcel->getActiveSheet()->setCellValue('F'.$i, '\''.$mqt_alsm_row[2])->getStyle('F'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->setCellValue('E'.$i, $MQT_ALSM_SUM_mqtname[$tmp])->getStyle('E'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->setCellValue('F'.$i, '\''.$mqt_alsm_row[2])->getStyle('F'.$i)->applyFromArray($styleArray01G);
			$objPHPExcel->getActiveSheet()->setCellValue('L'.$i, round(array_sum($MQT_ALSM_SUM_MQT_Unit_Qty_Prio[$tmp])))->getStyle('L'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			
			if (array_sum($MQT_ALSM_AA_FP[$tmp])>0)
			$OPD=(array_sum($MQT_ALSM_AA_FP[$tmp])-array_sum($MQT_ALSM_AA[$tmp]))/array_sum($MQT_ALSM_AA_FP[$tmp]);
			else $OPD=0;

			If ($OPD==0.00){
			$objPHPExcel->getActiveSheet()->getStyle('M'.$i)->getFont()->getColor()->setRGB('FF0000');
			$objPHPExcel->getActiveSheet()->getStyle('P'.$i)->getFont()->getColor()->setRGB('FF0000');
			}
			$M=array_sum($MQT_ALSM_AA_FP[$tmp])-(array_sum($MQT_ALSM_AA_FP[$tmp])*$OPD);

			$objPHPExcel->getActiveSheet()->setCellValue('M'.$i, $M)->getStyle('M'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->getStyle('M'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	

			$objPHPExcel->getActiveSheet()->setCellValue('N'.$i, array_sum($MQT_ALSM_AA_FP[$tmp]))->getStyle('N'.$i)->applyFromArray($styleArray01G);
			$objPHPExcel->getActiveSheet()->getStyle('N'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
			$objPHPExcel->getActiveSheet()->setCellValue('O'.$i, '')->getStyle('O'.$i)->applyFromArray($styleArray01G);
			$objPHPExcel->getActiveSheet()->getStyle('P'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
			$objPHPExcel->getActiveSheet()->setCellValue('P'.$i, $OPD)->getStyle('P'.$i)->applyFromArray($styleArray01G);
			$objPHPExcel->getActiveSheet()->setCellValue('Q'.$i, '-')->getStyle('Q'.$i)->applyFromArray($styleArray01G);
			
			$objPHPExcel->getActiveSheet()->setCellValue('R'.$i, round(array_sum($MQT_ALSM_AB[$tmp])))->getStyle('R'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->getStyle('R'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
			$objPHPExcel->getActiveSheet()->setCellValue('S'.$i, '-')->getStyle('S'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->getStyle('S'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
			$objPHPExcel->getActiveSheet()->setCellValue('T'.$i, round(array_sum($MQT_ALSM_AC[$tmp])))->getStyle('T'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->getStyle('T'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
			
			$objPHPExcel->getActiveSheet()->setCellValue('R'.$i, round(array_sum($MQT_ALSM_AB[$tmp])))->getStyle('R'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->getStyle('R'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
			$objPHPExcel->getActiveSheet()->setCellValue('S'.$i, '-')->getStyle('S'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->getStyle('S'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
			$objPHPExcel->getActiveSheet()->setCellValue('T'.$i, round(array_sum($MQT_ALSM_AC[$tmp])))->getStyle('T'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->getStyle('T'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	

			$U=array_sum($MQT_ALSM_Z_FP[$tmp])-(array_sum($MQT_ALSM_Z_FP[$tmp])*$OPD);
			$objPHPExcel->getActiveSheet()->setCellValue('U'.$i, $U)->getStyle('U'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);


			$objPHPExcel->getActiveSheet()->getStyle('U'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
			$objPHPExcel->getActiveSheet()->setCellValue('V'.$i, array_sum($MQT_ALSM_Z_FP[$tmp]))->getStyle('V'.$i)->applyFromArray($styleArray01);
			
			//$objPHPExcel->getActiveSheet()->setCellValue('W'.$i, round(array_sum($MQT_ALSM_BQ[$tmp])))->getStyle('W'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->setCellValue('W'.$i, '-')->getStyle('W'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			
			$objPHPExcel->getActiveSheet()->getStyle('W'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
			
			//$objPHPExcel->getActiveSheet()->setCellValue('X'.$i, '=Y'.$i.'-(Y'.$i.'*P'.$i.')')->getStyle('X'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->setCellValue('X'.$i, '-')->getStyle('X'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
						
			$objPHPExcel->getActiveSheet()->getStyle('X'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
			$objPHPExcel->getActiveSheet()->setCellValue('Y'.$i, array_sum($MQT_ALSM_Z_FP[$tmp]))->getStyle('Y'.$i)->applyFromArray($styleArray01);
			
			$objPHPExcel->getActiveSheet()->setCellValue('Z'.$i, '=AF'.$i.'-L'.$i.'')->getStyle('Z'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->getStyle('Z'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
			
			if ((array_sum($MQT_ALSM_AK[$tmp])-array_sum($MQT_ALSM_AA[$tmp]))>0)
			$OPD_AD=(((array_sum($MQT_ALSM_AK_FP[$tmp])-array_sum($MQT_ALSM_AA_FP[$tmp]))-(array_sum($MQT_ALSM_AK[$tmp])-array_sum($MQT_ALSM_AA[$tmp])))/(array_sum($MQT_ALSM_AK_FP[$tmp])-array_sum($MQT_ALSM_AA_FP[$tmp])));
			else $OPD_AD=0;	
			$AA=(array_sum($MQT_ALSM_AK_FP[$tmp])-array_sum($MQT_ALSM_AA_FP[$tmp]))-((array_sum($MQT_ALSM_AK_FP[$tmp])-array_sum($MQT_ALSM_AA_FP[$tmp]))*$OPD_AD);

			//$objPHPExcel->getActiveSheet()->setCellValue('AA'.$i, '=AB'.$i.'-(AB'.$i.'*AD'.$i.')')->getStyle('AA'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->setCellValue('AA'.$i, $AA)->getStyle('AA'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

			$objPHPExcel->getActiveSheet()->getStyle('AA'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
			$objPHPExcel->getActiveSheet()->setCellValue('AB'.$i, '=AH'.$i.'-N'.$i.'')->getStyle('AB'.$i)->applyFromArray($styleArray01G);
			$objPHPExcel->getActiveSheet()->setCellValue('AC'.$i, '-')->getStyle('AC'.$i)->applyFromArray($styleArray01G);
			//$objPHPExcel->getActiveSheet()->setCellValue('AD'.$i, '=((AH'.$i.'-N'.$i.')-(AG'.$i.'-M'.$i.'))/(AH'.$i.'-N'.$i.')')->getStyle('AD'.$i)->applyFromArray($styleArray01G);
			$objPHPExcel->getActiveSheet()->setCellValue('AD'.$i, $OPD_AD)->getStyle('AD'.$i)->applyFromArray($styleArray01G);
			$objPHPExcel->getActiveSheet()->getStyle('AD'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
			$objPHPExcel->getActiveSheet()->setCellValue('AE'.$i, '-')->getStyle('AE'.$i)->applyFromArray($styleArray01G);			
			
			
			
			$objPHPExcel->getActiveSheet()->setCellValue('AF'.$i, round(array_sum($MQT_ALSM_AN[$tmp])))->getStyle('AF'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->getStyle('AF'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
			
			if (array_sum($MQT_ALSM_AK[$tmp])>0)
			$OPD_AK=(array_sum($MQT_ALSM_AK_FP[$tmp])-array_sum($MQT_ALSM_AK[$tmp]))/array_sum($MQT_ALSM_AK_FP[$tmp]);
			else $OPD_AK=0;
			
			$AG=array_sum($MQT_ALSM_AK_FP[$tmp])-(array_sum($MQT_ALSM_AK_FP[$tmp])*$OPD_AK);
			//$objPHPExcel->getActiveSheet()->setCellValue('AG'.$i, '=AH'.$i.'-(AH'.$i.'*AJ'.$i.')')->getStyle('AG'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->setCellValue('AG'.$i, $AG)->getStyle('AG'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->getStyle('AG'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
			$objPHPExcel->getActiveSheet()->setCellValue('AH'.$i, array_sum($MQT_ALSM_AK_FP[$tmp]))->getStyle('AH'.$i)->applyFromArray($styleArray01);
			$objPHPExcel->getActiveSheet()->getStyle('AH'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);
			$objPHPExcel->getActiveSheet()->setCellValue('AI'.$i, '-')->getStyle('AI'.$i)->applyFromArray($styleArray01);
			$objPHPExcel->getActiveSheet()->setCellValue('AJ'.$i, $OPD_AK)->getStyle('AJ'.$i)->applyFromArray($styleArray01);
			$objPHPExcel->getActiveSheet()->getStyle('AJ'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
			$objPHPExcel->getActiveSheet()->setCellValue('AK'.$i, '-')->getStyle('AK'.$i)->applyFromArray($styleArray01);
			$objPHPExcel->getActiveSheet()->setCellValue('AL'.$i, round(array_sum($MQT_ALSM_AL[$tmp])))->getStyle('AL'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->getStyle('AL'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
			$objPHPExcel->getActiveSheet()->setCellValue('AM'.$i, '-')->getStyle('AM'.$i)->applyFromArray($styleArray01G)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->getStyle('AM'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
			$objPHPExcel->getActiveSheet()->setCellValue('AN'.$i, round(array_sum($MQT_ALSM_AN2[$tmp])))->getStyle('AN'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->getStyle('AN'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
			$objPHPExcel->getActiveSheet()->setCellValue('AO'.$i, array_sum($MQT_ALSM_AO[$tmp]))->getStyle('AO'.$i)->applyFromArray($styleArray01G);
			//$objPHPExcel->getActiveSheet()->setCellValue('AO'.$i, '=AP'.$i.'-(AP'.$i.'*AJ'.$i.')')->getStyle('AO'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->getStyle('AO'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
			$objPHPExcel->getActiveSheet()->setCellValue('AP'.$i, '-')->getStyle('AP'.$i)->applyFromArray($styleArray01);
			$objPHPExcel->getActiveSheet()->setCellValue('AQ'.$i, '-')->getStyle('AQ'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			$objPHPExcel->getActiveSheet()->getStyle('AQ'.$i)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);	
			$objPHPExcel->getActiveSheet()->setCellValue('AR'.$i, '-')->getStyle('AR'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			//$objPHPExcel->getActiveSheet()->setCellValue('AR'.$i, '=AS'.$i.'-(AS'.$i.'*AJ'.$i.')')->getStyle('AR'.$i)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
			
			$i++;
			$cnt++;
		}
	}


$cnt_s=$i;
$i++;
$objPHPExcel->getActiveSheet()->setCellValue('AT'.$i, 'Error Codes')->getStyle('AT'.$i)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$i++;
$objPHPExcel->getActiveSheet()->setCellValue('AT'.$i, 'E1:Prior IB Qty + Change in IB Qty does not equal Current IB Qty');
$i++;
$objPHPExcel->getActiveSheet()->setCellValue('AT'.$i, 'E2:Prior IB PSP + Change in IB PSP does not equal Current IB PSP');

If ($is_man_tr) {
$objPHPExcel->setActiveSheetIndex(2);
foreach (array_unique($Node_collector_m) as $tmp)
{
$Node_s_row=explode(":",$tmp);
$objPHPExcel->getActiveSheet()->setCellValue('B'.$i_ns, $cnt_ns)->getStyle('B'.$i_ns)->applyFromArray($styleArray01);
$objPHPExcel->getActiveSheet()->setCellValue('C'.$i_ns, 'Manual')->getStyle('C'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('D'.$i_ns, $Node_s_row[0])->getStyle('D'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('E'.$i_ns, $Node_s_row[1])->getStyle('E'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('F'.$i_ns, '')->getStyle('F'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('G'.$i_ns, '')->getStyle('G'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('H'.$i_ns, $Node_s_row[2])->getStyle('H'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('I'.$i_ns, $Node_s_row[3])->getStyle('I'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('J'.$i_ns, '\''.$Node_s_row[7])->getStyle('J'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('J'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->setCellValue('K'.$i_ns, '')->getStyle('K'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('L'.$i_ns, array_sum($Node_collector_m_nodes[$tmp]))->getStyle('L'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
//$objPHPExcel->getActiveSheet()->setCellValue('M'.$i_ns, '')->getStyle('M'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('N'.$i_ns, 'Not Available')->getStyle('N'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('N'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->setCellValue('O'.$i_ns, array_sum($Node_collector_m_mqtv[$tmp]))->getStyle('O'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('O'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
$objPHPExcel->getActiveSheet()->getStyle('O'.$i_ns)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('P'.$i_ns, $Node_s_row[4])->getStyle('P'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('Q'.$i_ns, '')->getStyle('Q'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('R'.$i_ns, array_sum($Node_collector_m_psp[$tmp]))->getStyle('R'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('R'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
$objPHPExcel->getActiveSheet()->getStyle('R'.$i_ns)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('S'.$i_ns, $Node_s_row[6]/100)->getStyle('S'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('S'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
$objPHPExcel->getActiveSheet()->setCellValue('T'.$i_ns, (1-($Node_s_row[6]/100))*array_sum($Node_collector_m_psp[$tmp]))->getStyle('T'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('T'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
$objPHPExcel->getActiveSheet()->getStyle('T'.$i_ns)->getFont()->getColor()->setRGB('0070C0');
$objPHPExcel->getActiveSheet()->setCellValue('U'.$i_ns, $Node_s_row[5])->getStyle('U'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('M'.$i_ns, '-')->getStyle('M'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$cnt_ns++;
$i_ns++;
}
}	
	if ($is_alsm_ok)
	{
		$objPHPExcel->setActiveSheetIndex(2);
		foreach (array_unique($Node_collector_alsm) as $tmp)
		{
			$Node_alsm_row=explode(":",$tmp);
			$objPHPExcel->getActiveSheet()->setCellValue('B'.$i_ns, $cnt_ns)->getStyle('B'.$i_ns)->applyFromArray($styleArray01);
			$objPHPExcel->getActiveSheet()->setCellValue('C'.$i_ns, 'ALSM')->getStyle('C'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->setCellValue('D'.$i_ns, $Node_alsm_row[0])->getStyle('D'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->setCellValue('E'.$i_ns, $Node_alsm_row[1])->getStyle('E'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->setCellValue('F'.$i_ns, '')->getStyle('F'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->setCellValue('G'.$i_ns, '')->getStyle('G'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->setCellValue('H'.$i_ns, $Node_collector_alsm_mqt_prod[$tmp][0])->getStyle('H'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->setCellValue('I'.$i_ns, $Node_collector_alsm_mqt_config[$tmp][0])->getStyle('I'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->setCellValue('J'.$i_ns, '\''.$Node_alsm_row[2])->getStyle('J'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			//$objPHPExcel->getActiveSheet()->setCellValue('L'.$i_ns, count(array_unique($Node_collector_alsm_nodes[$tmp])))->getStyle('L'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->setCellValue('L'.$i_ns, $Node_collector_alsm_nodes[$tmp])->getStyle('L'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);			
			$objPHPExcel->getActiveSheet()->setCellValue('M'.$i_ns, '-')->getStyle('M'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->setCellValue('N'.$i_ns, array_sum($Node_collector_alsm_q_parts[$tmp]))->getStyle('N'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			//$objPHPExcel->getActiveSheet()->getStyle('N'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
			$objPHPExcel->getActiveSheet()->setCellValue('O'.$i_ns, array_sum($Node_collector_alsm_mqtvaltotal[$tmp]))->getStyle('O'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			//$objPHPExcel->getActiveSheet()->getStyle('O'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER);
			$objPHPExcel->getActiveSheet()->setCellValue('P'.$i_ns, $Node_collector_alsm_mqtname[$tmp][0])->getStyle('P'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->setCellValue('Q'.$i_ns, '')->getStyle('Q'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->setCellValue('R'.$i_ns, array_sum($Node_collector_alsm_glp_price_total[$tmp]))->getStyle('R'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->getStyle('R'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
			//$objPHPExcel->getActiveSheet()->getStyle('R'.$i_ns)->getFont()->getColor()->setRGB('0070C0');
			$objPHPExcel->getActiveSheet()->setCellValue('S'.$i_ns, $Node_collector_alsm_disc[$tmp][0])->getStyle('S'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->getStyle('S'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00);
			$objPHPExcel->getActiveSheet()->setCellValue('T'.$i_ns, array_sum($Node_collector_alsm_psp_price_total[$tmp]))->getStyle('T'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
			$objPHPExcel->getActiveSheet()->getStyle('T'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_ACCOUNTING);	
			//$objPHPExcel->getActiveSheet()->getStyle('T'.$i_ns)->getFont()->getColor()->setRGB('0070C0');
			
			$cnt_ns++;
			$i_ns++;
		}
	}


$objPHPExcel->getActiveSheet()->setCellValue('L'.$i_ns, '=SUBTOTAL(9,L8:L'.($i_ns-1).')')->getStyle('L'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('M'.$i_ns, '=SUBTOTAL(9,M8:M'.($i_ns-1).')')->getStyle('M'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('N'.$i_ns, '=SUBTOTAL(9,N8:N'.($i_ns-1).')')->getStyle('N'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('O'.$i_ns, '=SUBTOTAL(9,O8:O'.($i_ns-1).')')->getStyle('O'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->setCellValue('R'.$i_ns, '=SUBTOTAL(9,R8:R'.($i_ns-1).')')->getStyle('R'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->setCellValue('T'.$i_ns, '=SUBTOTAL(9,T8:T'.($i_ns-1).')')->getStyle('T'.$i_ns)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

If ($runc['curr']=='EUR') 
{
	$objPHPExcel->getActiveSheet()->getStyle('R'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('T'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	}
else 
{
	$objPHPExcel->getActiveSheet()->getStyle('R'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);	
	$objPHPExcel->getActiveSheet()->getStyle('T'.$i_ns)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
}

	$objPHPExcel->setActiveSheetIndex(1);
	$objPHPExcel->getActiveSheet()->setCellValue('L'.$cnt_s, '=SUBTOTAL(9,L8:L'.($cnt_s-1).')')->getStyle('L'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('M'.$cnt_s, '=SUBTOTAL(9,M8:M'.($cnt_s-1).')')->getStyle('M'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('N'.$cnt_s, '=SUBTOTAL(9,N8:N'.($cnt_s-1).')')->getStyle('N'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('R'.$cnt_s, '=SUBTOTAL(9,R8:R'.($cnt_s-1).')')->getStyle('R'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('S'.$cnt_s, '=SUBTOTAL(9,S8:S'.($cnt_s-1).')')->getStyle('S'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('T'.$cnt_s, '=SUBTOTAL(9,T8:T'.($cnt_s-1).')')->getStyle('T'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('U'.$cnt_s, '=SUBTOTAL(9,U8:U'.($cnt_s-1).')')->getStyle('U'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('W'.$cnt_s, '=SUBTOTAL(9,W8:W'.($cnt_s-1).')')->getStyle('W'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('X'.$cnt_s, '=SUBTOTAL(9,X8:X'.($cnt_s-1).')')->getStyle('X'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('Z'.$cnt_s, '=SUBTOTAL(9,Z8:Z'.($cnt_s-1).')')->getStyle('Z'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('AA'.$cnt_s, '=SUBTOTAL(9,AA8:AA'.($cnt_s-1).')')->getStyle('AA'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('AB'.$cnt_s, '=SUBTOTAL(9,AB8:AB'.($cnt_s-1).')')->getStyle('AB'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('AF'.$cnt_s, '=SUBTOTAL(9,AF8:AF'.($cnt_s-1).')')->getStyle('AF'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('AG'.$cnt_s, '=SUBTOTAL(9,AG8:AG'.($cnt_s-1).')')->getStyle('AG'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('AH'.$cnt_s, '=SUBTOTAL(9,AH8:AH'.($cnt_s-1).')')->getStyle('AH'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('AL'.$cnt_s, '=SUBTOTAL(9,AL8:AL'.($cnt_s-1).')')->getStyle('AL'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('AM'.$cnt_s, '=SUBTOTAL(9,AM8:AM'.($cnt_s-1).')')->getStyle('AM'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('AN'.$cnt_s, '=SUBTOTAL(9,AN8:AN'.($cnt_s-1).')')->getStyle('AN'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('AO'.$cnt_s, '=SUBTOTAL(9,AO8:AO'.($cnt_s-1).')')->getStyle('AO'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('AQ'.$cnt_s, '=SUBTOTAL(9,AQ8:AQ'.($cnt_s-1).')')->getStyle('AQ'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
	$objPHPExcel->getActiveSheet()->setCellValue('AR'.$cnt_s, '=SUBTOTAL(9,AR8:AR'.($cnt_s-1).')')->getStyle('AR'.$cnt_s)->applyFromArray($styleArray01)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);


If ($runc['curr']=='EUR') 
{
	$objPHPExcel->getActiveSheet()->getStyle('M'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('N'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('U'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	//$objPHPExcel->getActiveSheet()->getStyle('T'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('X'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AA'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AB'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AG'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AH'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AO'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	$objPHPExcel->getActiveSheet()->getStyle('AR'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
	
}
else 
{
	$objPHPExcel->getActiveSheet()->getStyle('M'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);	
	$objPHPExcel->getActiveSheet()->getStyle('N'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('U'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	//$objPHPExcel->getActiveSheet()->getStyle('T'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('X'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AA'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);	
	$objPHPExcel->getActiveSheet()->getStyle('AB'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AG'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);	
	$objPHPExcel->getActiveSheet()->getStyle('AH'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
	$objPHPExcel->getActiveSheet()->getStyle('AO'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);	
	$objPHPExcel->getActiveSheet()->getStyle('AR'.$cnt_s)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD);
}	
//print_r(array_unique($Node_collector));
//print_r($Node_collector_m);
//print_r($Node_collector_nbofnodes);

$no_ts_unit=array_count_values($no_ts_unit);
$notsscheck=sizeof($no_ts_unit);
$i1=$cnt_mts;

if ($notsscheck)
{
	$objPHPExcel->setActiveSheetIndex(0);
	
	$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, "Table C")->getStyle('B'.$i1)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFont()->setBold( true );
	$i1=$i1+1;
	$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, "Part Number (unidentified)")->getStyle('B'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFont()->setBold( true );
	$objPHPExcel->getActiveSheet()->setCellValue('C'.$i1, "Amount of cards")->getStyle('C'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('C'.$i1)->getFont()->setBold( true );

		foreach($no_ts_unit as $key => $value) 
		{
			$i1++;
			$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, $key)->getStyle('B'.$i1)->applyFromArray($styleArray01);
			$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFont()->getColor()->setRGB('FF0000');
			$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFont()->setBold( true );
			$objPHPExcel->getActiveSheet()->setCellValue('C'.$i1, $value)->getStyle('C'.$i1)->applyFromArray($styleArray01);
		}
	
	$i1=$i1+2;
	
	$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, "This CRA is incomplete,\nunsupported Part Number\ndetected. Data was\ndetected in table C that\ncould not be\nprocessed.")->getStyle('B'.$i1)->applyFromArray($styleArray01);	
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFont()->setBold( true );
	//$objPHPExcel->getActiveSheet()->getRowDimension($i1)->setRowHeight(45);
	$objPHPExcel->getActiveSheet()->setCellValue('C'.$i1, "Contact the Sonar\nAdmin or the Care PLM\nresponsible for the\nproduct")->getStyle('C'.$i1)->applyFromArray($styleArray01);
	$objPHPExcel->getActiveSheet()->getStyle('C'.$i1)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->getStyle('C'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');	
	$objPHPExcel->getActiveSheet()->setCellValue('D'.$i1, "Sonar Admin")->getStyle('D'.$i1) ->applyFromArray($styleArray01_L);
	$url="http://ddtb.de.alcatel-lucent.com/cib/support.php";
	$objPHPExcel->getActiveSheet()->getHyperlink('D'.$i1)->setUrl($url);	
	$objPHPExcel->getActiveSheet()->getStyle('D'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');	
	$objPHPExcel->getActiveSheet()->setCellValue('E'.$i1, "Care PLM Prime")->getStyle('E'.$i1)->applyFromArray($styleArray01_L);
	$objPHPExcel->getActiveSheet()->getStyle('E'.$i1)->getFill()->getStartColor()->setRGB('F2F2F2');
	$url="http://nok.it/CarePLMPrime";
	$objPHPExcel->getActiveSheet()->getHyperlink('E'.$i1)->setUrl($url);	
	$redA=1;
}

if ($alsm_not_taged)
{
	$i1=$i1+2;
	$objPHPExcel->getActiveSheet()->setCellValue('B'.$i1, "Please review ALSM data, assign it the NTAG")->getStyle('B'.$i1)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFont()->getColor()->setRGB('FF0000');
	$objPHPExcel->getActiveSheet()->getStyle('B'.$i1)->getFont()->setBold( true );
	$redA=1;
}

If ($redA==1){
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->setCellValue('A12', "This CRA is INCOMPLETE or contains ERRORS and should not be used.  Please take appropriate Next Steps to correct the issue and re-run the CRA.")->getStyle('A12')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$objPHPExcel->getActiveSheet()->getStyle('A12')->getFont()->getColor()->setRGB('FF0000');
$objPHPExcel->getActiveSheet()->getStyle('A12')->getFont()->setBold( true );
}

if ($HideTab){

	foreach ($tabsnames as $value)
	{
	$objPHPExcel->getSheetByName($value)->setSheetState(PHPExcel_Worksheet::SHEETSTATE_HIDDEN);
	}
}
/*
$objPHPExcel->removeSheetByIndex(3);
$objPHPExcel->removeSheetByIndex(2);
$objPHPExcel->removeSheetByIndex(1);
*/

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$year = (new DateTime)->format("Y");
$xlsfilnameS='/mqt/'.$year.'/'.$xlsfilname;
$objWriter->save($save_path.$xlsfilnameS);
cib_dbQuery("TRUNCATE `mqt_rep`");
cib_dbQuery("update `mqt_rep_request` SET `when_compl`= NOW(),`run` = '1',`file_name` = '".$xlsfilname."' where `id`='".cib_dbEscape($runc['id'])."' LIMIT 1");
cib_dbQuery("update `mailcronrun` SET `mqt_rep`='0' LIMIT 1");
cib_dbQuery("UPDATE `mqt_rep_request` SET `approve` = '0' WHERE `cuid`='".cib_dbEscape($runc['cuid'])."' AND `opnbr`='".cib_dbEscape($runc['opnbr'])."' ");	

$dlink='<a href="'.$cib_settings['site_url'].'download_file_simple.php?name=.'.$xlsfilnameS.'" target="blank" style="text-decoration: none;">'.$xlsfilname.'</a>';		
						
$msubject="Request MQT Report completed";
$message="Customer name: ".$runc['cuname'].
"<br>Country: ".$runc['cname'].
"<br>sONAr ID: ".$runc['cuid'].
"<br>NTAGs: ".$ntags_f.
"<br>Link to the file: ".$dlink;

$headers = "MIME-Version: 1.0\n" ; 
$headers .= "Content-Type: text/html; charset=\"iso-8859-1\"\n"; 
$headers.="From: $cib_settings[support_mail]\n";
$headers.="Reply-to: $cib_settings[support_mail]\n";
$headers.="Bcc: pszymanski@alcatel-lucent.com\n";	
if ($cib_settings['smail']) @mail($runc['who_mail'],$msubject,$message,$headers);
?>
