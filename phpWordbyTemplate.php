<?php

require 'vendor/autoload.php';
use PhpOffice\PhpWord\Shared\Converter;

set_time_limit(3600);
// 現在時間
$t = time();
$time_now = date('Ymd_His', $t); // filename

// 載入範本文件
$tpl = new \PhpOffice\PhpWord\TemplateProcessor('tpl/word_tpl.docx');

// 設定值
$tpl->setValues(array('firstname' => 'John', 'lastname' => 'Doe'));

// 設定圖像值
$tpl->setValue('Name', 'John Doe');
$tpl->setValue(array('City', 'Street'), array('Detroit', '12th Street'));
$tpl->setImageValue('CompanyLogo', 'images/logo_irt.png');
$tpl->setImageValue('UserLogo', array('path' => 'images/user.jpg', 'width' => 100, 'height' => 100, 'ratio' => false));


// 複製區塊（方法1） block_name#1 block_name#2 block_name#3
// $tpl->cloneBlock('block_name', 3, true, true);

// 複製區塊（方法2）
$replacements = array(
    array('customer_name' => 'Batman', 'customer_address' => 'Gotham City'),
    array('customer_name' => 'Superman', 'customer_address' => 'Metropolis'),
);
$tpl->cloneBlock('block_name', 0, true, false, $replacements);

// ????? 替換區塊内容 ?????
$tpl->replaceBlock('block_replace', 'This is the replacement text.');

// 刪除區塊
$tpl->deleteBlock('block_delete');

// 表格複製行和設定值
$t1_values = [
    ['t1_num' => 1, 't1_name' => 'Batman', 't1_asum' => '33', 't1_arate' => '33%', 't1_bsum' => '66', 't1_brate' => '66%'],
    ['t1_num' => 2, 't1_name' => 'Superman', 't1_asum' => '22', 't1_arate' => '22%', 't1_bsum' => '77', 't1_brate' => '77%'],
];
$tpl->cloneRowAndSetValues('t1_num', $t1_values);
$tpl->setValues(array('t1_asum_tt' => '55', 't1_arate_tt' => '55%', 't1_bsum_tt' => '143', 't1_brate_tt' => '143%'));



// 設定圖表值
$categories = array('A', 'B', 'C', 'D', 'E');
$series1 = array(1, 3, 2, 5, 4);
$chart = new PhpOffice\PhpWord\Element\Chart('doughnut', $categories, $series1);
$chart->getStyle()
        ->setWidth(Converter::cmToEmu(16))
        ->setHeight(Converter::cmToEmu(5));

$tpl->setChart('myChart', $chart);

// 設定巨集開頭字符
// $tpl->setMacroOpeningChars('{#');
// 設定巨集關閉字符
// $tpl->setMacroClosingChars('#}');

$pathToSave = 'export/export.docx';
// $filename_download = 'hello_' . $time_now . '.docx';

$tpl->saveAs($pathToSave);

