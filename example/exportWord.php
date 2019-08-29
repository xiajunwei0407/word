<?php

// 调用示例
// 示例中的获取数据方式是基于Thinkphp5，请根据实际情况调整

function exportWord()
{
    $list = $this->model
        ->alias('a')
        ->join('tb_admin b', 'a.user_id = b.id', 'left')
        ->field('a.*,b.username')
        ->order('a.id', 'desc')
        ->select();
    $list = collection($list)->toArray();

    $html = '<table bgcolor="#000000" align="center">';
    $html .= '<tr bgcolor="white">';
    $html .= '<th style="border:1px solid #c8c8c8;">视频分类id</th>';
    $html .= '<th style="border:1px solid #c8c8c8;">视频分类</th>';
    $html .= '<th style="border:1px solid #c8c8c8;">关联标签</th>';
    $html .= '<th style="border:1px solid #c8c8c8;">创建人</th>';
    $html .= '<th style="border:1px solid #c8c8c8;">创建时间</th>';
    $html .= '</tr>';
    foreach ($list as $key => $val) {
        $html .= '<tr bgcolor="white">';
        $html .= '<td style="border:1px solid #c8c8c8;">'. $val['id'] .'</td>';
        $html .= '<td style="border:1px solid #c8c8c8;">'. $val['name'] .'</td>';
        $html .= '<td style="border:1px solid #c8c8c8;">'. $val['label_id_text'] .'</td>';
        $html .= '<td style="border:1px solid #c8c8c8;">'. $val['username'] .'</td>';
        $html .= '<td style="border:1px solid #c8c8c8;">'. $val['create_time_text'] .'</td>';
        $html .= '</tr>';
    }
    $html .= '</table>';
    ///////////////////////////////////////
    $word = new \Xjw\Word();
    $word->start();
    $fileName = '视频分类.doc';
    echo $html;
    $word->save($fileName);
    ///////////////////////////////////////
    // 文件的类型
    header('Content-type: application/word');
    header("Content-Disposition: attachment; filename=$fileName");
    readfile($fileName);
    ob_flush();
    flush();
    exit();
}