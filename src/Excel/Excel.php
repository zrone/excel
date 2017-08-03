<?php
/**
 * Created by PhpStorm.
 * User: zrone
 * Date: 16/3/23
 * Time: 16:40
 */

/**
 * Class Excel
 *
 * © zrone <xujining2008@126.com>
 *
 * @package Excel
 */
class Excel
{
    private $_data;
    private $_name;
    private $_cellName;
    private $instance = null;
    private $_fileArray = array ();

    /**
     * Excel constructor.
     *
     * @param array $data
     * @param       $name
     * @param       $cellName
     */
    public function __construct( array $data, $name, $cellName, $fileArray = array () )
    {
        if ( !is_array( $data ) ) {
            return false;
        }
        require_once( dirname( __FILE__ ) . DIRECTORY_SEPARATOR . "Src/PHPExcel.php" );
        require_once dirname( __FILE__ ) . DIRECTORY_SEPARATOR . "Src/PHPExcel/Worksheet/Drawing.php";
        $this->instance = new PHPExcel();
        $this->instance->setActiveSheetIndex( 0 );
        $this->_data = $data;
        $this->_name = trim( $name );
        $this->_cellName = $cellName;
        $this->_fileArray = $fileArray;
    }

    /**
     * @return array|bool
     */
    private function handleCellTitile()
    {
        $cellArr = explode( "|", $this->_cellName );
        /** @var $cellKey array 对应数据.类型 */
        $cellKey = array ();
        /** @var $cellValue array 列名称 */
        $cellValue = array ();
        /** @var $number integer 列数 */
        $number = count( $cellArr );
        if ( $number > 0 ) {
            for ( $i = 0; $i < $number; $i ++ ) {
                $temp = explode( "::", $cellArr[ $i ] );
                $cellKey[] = $temp[ 0 ];
                $cellValue[] = $temp[ 1 ];
            }

            return array (
                $cellKey,
                $cellValue
            );
        }

        return false;
    }

    /**
     * @param $key
     *
     * @return array|bool
     */
    private function handleKey( $key )
    {
        $key = trim( $key );
        $keyObj = preg_match( "/\{.*\}/", $key, $temp );
        if ( $keyObj ) {
            /** @var $realKey string 数据字段 */
            $realKey = substr( $key, 0, strpos( $key, "{" ) );

            return array (
                'realKey' => $realKey,
                'temp'    => $temp[ 0 ]
            );
        }

        return false;
    }

    /**
     * @return PHPExcel|null
     */
    private function handleData()
    {
        $arrayObj = new \arrayObject( $this->_data );
        $iterator = $arrayObj->getIterator();
        $index = 1;
        $cellHeader = $this->handleCellTitile();
        $cellKey = $cellHeader[ 0 ];
        $cellTitle = $cellHeader[ 1 ];
        for ( $ascii = 65, $i = 0; $i < count( $cellTitle ); $i ++ ) {
            $this->instance->setActiveSheetIndex( 0 )
                ->setCellValue( chr( $ascii ) . $index, $cellTitle[ $i ] )
                ->getDefaultStyle()
                ->getAlignment()
                ->setHorizontal( PHPExcel_Style_Alignment::HORIZONTAL_CENTER )
                ->setVertical( PHPExcel_Style_Alignment::VERTICAL_CENTER );
            $this->instance->getActiveSheet()
                ->getColumnDimension( chr( $ascii ) )
                ->setWidth( 25 );
            $ascii ++;
        }
        // 标题设置背景
        $this->instance->getActiveSheet()
            ->getStyle( chr( 65 ) . '1:' . chr( $ascii - 1 ) . '1' )
            ->getFill()
            ->setFillType( PHPExcel_Style_Fill::FILL_SOLID )
            ->getStartColor()
            ->setRGB( 'C6EFCE' );
        // 设置标题边框
        $this->instance->getActiveSheet()
            ->getStyle( chr( 65 ) . '1:' . chr( $ascii - 1 ) . '1' )
            ->getBorders()
            ->getBottom()
            ->setBorderStyle( PHPExcel_Style_Border::BORDER_MEDIUM );

        $this->instance->getActiveSheet()
            ->getRowDimension( 1 )
            ->setRowHeight( 55 );
        $iterator->rewind();
        while ( $iterator->valid() ) {
            $index ++;
            $currentValue = $iterator->current();
            for ( $ascii = 65, $j = 0; $j < count( $cellKey ); $j ++ ) {
                $temp = $this->handleKey( $cellKey[ $j ] );
                if ( $temp ) {
                    /** @var $tempOptions array */
                    $tempOptions = json_decode( $temp[ 'temp' ], true );
                    switch ( $tempOptions[ 'element_type' ] ) {
                        case 'string':
                            $this->instance->setActiveSheetIndex( 0 )
                                ->setCellValueExplicit( chr( $ascii ) . $index, $currentValue[ $temp[ 'realKey' ] ], PHPExcel_Cell_DataType::TYPE_STRING )
                                ->getDefaultStyle()
                                ->getAlignment()
                                ->setHorizontal( PHPExcel_Style_Alignment::HORIZONTAL_CENTER )
                                ->setVertical( PHPExcel_Style_Alignment::VERTICAL_CENTER );
                            break;
                        case 'radio':
                            $this->instance->setActiveSheetIndex( 0 )
                                ->setCellValueExplicit( chr( $ascii ) . $index, $tempOptions[ $currentValue[ $temp[ 'realKey' ] ] ], PHPExcel_Cell_DataType::TYPE_STRING )
                                ->getDefaultStyle()
                                ->getAlignment()
                                ->setHorizontal( PHPExcel_Style_Alignment::HORIZONTAL_CENTER )
                                ->setVertical( PHPExcel_Style_Alignment::VERTICAL_CENTER );
                            break;
                        case 'image':
                            $this->instance->setActiveSheetIndex( 0 );
                            $this->instance->getActiveSheet()
                                ->getColumnDimension( chr( $ascii ) )
                                ->setWidth( 10 );

                            if ( is_file( $currentValue[ $temp[ 'realKey' ] ] ) ) {

                                $draw = new PHPExcel_Worksheet_Drawing();
                                $draw->setPath( $currentValue[ $temp[ 'realKey' ] ] )
                                    ->setResizeProportional( false )
                                    ->setOffsetY( 10 )
                                    ->setOffsetX( 10 )
                                    ->setWidth( 50 )
                                    ->setHeight( 50 )
                                    ->setCoordinates( chr( $ascii ) . $index )
                                    ->setRotation( 0 )
                                    ->getShadow()
                                    ->setVisible( true )
                                    ->setDirection( 75 );
                                $draw->setWorksheet( $this->instance->getActiveSheet() );
                                unset( $draw );
                            } else {
                                $this->instance->setActiveSheetIndex( 0 )
                                    ->setCellValueExplicit( chr( $ascii ) . $index, "", PHPExcel_Cell_DataType::TYPE_STRING )
                                    ->getDefaultStyle()
                                    ->getAlignment()
                                    ->setHorizontal( PHPExcel_Style_Alignment::HORIZONTAL_CENTER )
                                    ->setVertical( PHPExcel_Style_Alignment::VERTICAL_CENTER );
                            }
                            break;
                    }
                }
                $ascii ++;
            }
            $this->instance->getActiveSheet()
                ->getRowDimension( $index )
                ->setRowHeight( 55 );
            $iterator->next();
        }

        return $this->instance;
    }

    /**
     * export
     *
     * @return void
     */
    public function export()
    {
        $instance = $this->handleData();
        $instance->getActiveSheet()
            ->setTitle( $this->_name );

        $instance->setActiveSheetIndex( 0 );
        header( 'Content-Type: application/vnd.ms-excel' );
        header( 'Content-Disposition: attachment;filename="' . $this->_name . '.xlsx"' );
        header( 'Cache-Control: max-age=0' );
        $objWriter = PHPExcel_IOFactory::createWriter( $instance, 'Excel2007' );
        $objWriter->save( 'php://output' );
        if ( $this->_fileArray ) {
            for ( $i = 0; $i < count( $this->_fileArray ); $i ++ ) {
                @unlink( $this->_fileArray[ $i ] );
            }
        }
    }

    /**
     * web image to file
     *
     * @param $url
     *
     * @return null|string
     */
    public function getWebImage( $url )
    {
        if ( empty( $url ) ) {
            return "";
        }

        $filename = ROOT . DIRECTORY_SEPARATOR . 'upload' . DIRECTORY_SEPARATOR . str_shuffle( date( "Ymdhis" ) . rand( 1000, 9999 ) );

        $ch = curl_init();
        curl_setopt( $ch, CURLOPT_URL, $url );
        curl_setopt( $ch, CURLOPT_FOLLOWLOCATION, 1 );
        curl_setopt( $ch, CURLOPT_RETURNTRANSFER, 1 );
        curl_setopt( $ch, CURLOPT_CONNECTTIMEOUT, 30 );
        ob_start();
        curl_exec( $ch );
        $imageData = ob_get_contents();
        ob_end_clean();
        file_put_contents( $filename, $imageData );

        if ( is_file( $filename ) ) {
            $mime = finfo_file( finfo_open( FILEINFO_MIME_TYPE ), $filename );

            switch ( $mime ) {
                case 'image/jpeg':
                    $suffix = ".jpg";
                    break;
                case 'image/bmp':
                    $suffix = ".bmp";
                    break;
                case 'image/gif':
                    $suffix = ".gif";
                    break;
                case 'image/png':
                    $suffix = ".png";
                    break;
                default:
                    $suffix = ".jpg";
                    break;
            }
            @rename( $filename, $filename . $suffix );
            @chmod( $filename . $suffix, 777 );
            @unlink( $filename );
            unset( $imageData );

            return $filename . $suffix;
        } else {
            return "";
        }
    }
}