<?php
require_once './PHPExcel/IOFactory.php';
function getaddr($address,$type){
    if($address=='' || $address==null){
        return "读取地址失败！";
    }
    $prepAddr = str_replace(' ','+',$address);  
    $lat = $lng = 0;
    $geocode=file_get_contents('http://maps.google.com/maps/api/geocode/json?address='.$prepAddr.'&sensor=false');//如果部署到国内服务器需要把com改为cn
      
    $output= json_decode($geocode);  
    if($output->status=="OK"){
      $lat = $output->results[0]->geometry->location->lat;
      $lng = $output->results[0]->geometry->location->lng;
      if($type=="lat"){
        return $lat . ";" . $lng;
      }else{
        return $address ."的坐标是[北纬：" . $lat . "  东经：" . $lng . "]";
      }
    }else{
        $errmsg = $output->status;
        return $errmsg;
    }

}
// 检查excel文件是否存在
if (!file_exists("data.xls")) {
  exit("not found data.xls.\n");
}
//载入excel文件
$reader = PHPExcel_IOFactory::createReader('Excel5'); //设置以Excel5格式(Excel97-2003工作簿)
$PHPExcel = $reader->load("data.xls"); // 载入excel文件
$sheet = $PHPExcel->getSheet(0); // 读取第一個工作表
$highestRow = $sheet->getHighestRow(); // 取得总行数
$highestColumm = $sheet->getHighestColumn(); // 取得总列数

//
$file_name = 'location_' . time() . '.csv';
$data = '地名,北纬;东经'. "\n";

/** 循环读取并更新每个单元格的数据 */
for ($row = 2; $row <= $highestRow; $row++){//行数是以第2行开始
  for ($column = 'A'; $column <= $highestColumm; $column++) {//列数是以A列开始
    $dataset[] = $sheet->getCell($column.$row)->getValue();
    if($column=='A'){
      $data .= $sheet->getCell($column.$row)->getValue();
    }elseif($column=='B'){
      $data .= "," . getaddr($sheet->getCell('A'.$row)->getValue(),"lat");
    }
  }
  $data .= "\n";
}
//写入文件
$fp = fopen($file_name,"a");//打开输出文件，如果不存在则自动创建。
fwrite($fp,iconv('utf-8','gb2312//TRANSLIT//IGNORE', $data)); //写入数据
fclose($fp); //关闭文件句柄
echo "生成成功";
?>