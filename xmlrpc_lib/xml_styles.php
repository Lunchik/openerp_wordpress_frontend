<?php

    $style_normal = array(
                    'font' => array(
                        'name' => 'Open Sans',
                        'size' => 8,
                        'color' => array('argb' => 'FF036B7E')
                    )
        );

    $style_header = array(
                    'font' => array(
                        'name' => 'Open Sans',
                        'size' => 8,
                        'bold' => true,
                        'color' => array('argb' => 'FF036B7E')
                    ),
                    'fill' => array(
                        'type' => PHPExcel_Style_Fill::FILL_SOLID,
                        'color' => array('argb' => 'FFDCEFF2')
                    )
        );
    //total = header here
    $style_total = array(
                    'font' => array(
                        'name' => 'Open Sans',
                        'size' => 8,
                        'bold' => true,
                        'color' => array('argb' => 'FF036B7E')
                    ),
                    'fill' => array(
                        'type' => PHPExcel_Style_Fill::FILL_SOLID,
                        'color' => array('argb' => 'FFDCEFF2')
                    )
        );

    $style_ctitle = array(
                    'font' => array(
                        'name' => 'Open Sans',
                        'size' => 8,
                        'bold' => true,
                        'color' => array('argb' => 'FF036B7E')
                    )/*,
                    'fill' => array(
                        'type' => PHPExcel_Style_Fill::FILL_SOLID,
                        'color' => array('argb' => 'FFDCEFF2')
                    )*/
        );
    //ctitle = stotal here
    $style_stotal = array(
                    'font' => array(
                        'name' => 'Open Sans',
                        'size' => 8,
                        'bold' => true,
                        'color' => array('argb' => 'FF036B7E')
                    ),
                    'fill' => array(
                        'type' => PHPExcel_Style_Fill::FILL_SOLID,
                        'color' => array('argb' => 'FFDCEFF2')
                    )
        );

    $style_borders = array(
        'borders' => array(
            'allborders' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => array('argb' => 'FF036B7E')
            )
        )
    );

?>