<?php
/*
  Plugin Name: Simple PHPExcel Export
  Description: Simple PHPExcel Export Plugin for WordPress
  Version: 1.0.0
  Author: Mithun
  Author URI: http://twitter.com/mithunp
 */

  define("SPEE_PLUGIN_URL", WP_PLUGIN_URL.'/'.basename(dirname(__FILE__)));
  define("SPEE_PLUGIN_DIR", WP_PLUGIN_DIR.'/'.basename(dirname(__FILE__)));

  add_action ( 'admin_menu', 'spee_admin_menu' );

  function spee_admin_menu() {
  	add_menu_page ( 'PHPExcel Export', 'Export', 'manage_options', 'spee-dashboard', 'spee_dashboard' );
  }

  function spee_dashboard() {
  	global $wpdb;
  	if ( isset( $_GET['export'] )) {
  		if ( file_exists(SPEE_PLUGIN_DIR . '/lib/PHPExcel.php') ) {

			//Include PHPExcel
  			require_once (SPEE_PLUGIN_DIR . "/lib/PHPExcel.php");

			// Create new PHPExcel object
  			$objPHPExcel = new PHPExcel();

			// Set document properties

			// Add some data
  			$objPHPExcel->setActiveSheetIndex(0);
  			$objPHPExcel->getActiveSheet()->setCellValue('A1', 'No');
  			$objPHPExcel->getActiveSheet()->setCellValue('B1', 'ID');
  			$objPHPExcel->getActiveSheet()->setCellValue('C1', 'Author');
  			$objPHPExcel->getActiveSheet()->setCellValue('D1', 'Date');
  			$objPHPExcel->getActiveSheet()->setCellValue('E1', 'Title');
  			$objPHPExcel->getActiveSheet()->setCellValue('F1', 'Status');
  			$objPHPExcel->getActiveSheet()->setCellValue('G1', 'Content');
  			$objPHPExcel->getActiveSheet()->setCellValue('F1', 'Comment Count');

  			$objPHPExcel->getActiveSheet()->getStyle('A1:G1')->getFont()->setBold(true);
  			$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('A:G')->setAutoSize(true);

  			$author = $_GET['author'];
  			$start_date = $_GET['start_date'];
  			$end_date = $_GET['end_date'];

  			$query = "SELECT p.*, u.display_name
  			FROM {$wpdb->prefix}posts AS p
  			LEFT JOIN {$wpdb->prefix}users AS u ON p.post_author = u.ID
  			WHERE p.post_type = 'post'
  			AND p.post_status = 'publish'
  			AND u.display_name = '$author'
  			AND p.post_date BETWEEN '$start_date' AND '$end_date 23:59:59'";

  			$posts   = $wpdb->get_results($query);

  			if ( $posts ) {
  				foreach ( $posts as $i=>$post ) {
  					$objPHPExcel->getActiveSheet()->setCellValue('A'.($i+2), $i+1);
  					$objPHPExcel->getActiveSheet()->setCellValue('B'.($i+2), $post->ID);
  					$objPHPExcel->getActiveSheet()->setCellValue('C'.($i+2), $post->display_name);
  					$objPHPExcel->getActiveSheet()->setCellValue('D'.($i+2), $post->post_date);
  					$objPHPExcel->getActiveSheet()->setCellValue('E'.($i+2), $post->post_title);
  					$objPHPExcel->getActiveSheet()->setCellValue('F'.($i+2), $post->post_status);
  					$objPHPExcel->getActiveSheet()->setCellValue('G'.($i+2), $post->post_content);
  					$objPHPExcel->getActiveSheet()->setCellValue('F'.($i+2), $post->comment_count);
  				}
  			}

			// Rename worksheet
			//$objPHPExcel->getActiveSheet()->setTitle('Simple');

			// Set active sheet index to the first sheet, so Excel opens this as the first sheet
  			$objPHPExcel->setActiveSheetIndex(0);

			// Redirect output to a client’s web browser
  			ob_clean();
  			ob_start();
  			switch ( $_GET['format'] ) {
  				case 'csv':
					// Redirect output to a client’s web browser (CSV)
  				header("Content-type: text/csv");
  				header("Cache-Control: no-store, no-cache");
  				header('Content-Disposition: attachment; filename="Report'." ".$end_date." ".$post->display_name.'.csv"');
  				$objWriter = new PHPExcel_Writer_CSV($objPHPExcel);
  				$objWriter->setDelimiter(',');
  				$objWriter->setEnclosure('"');
  				$objWriter->setLineEnding("\r\n");
					//$objWriter->setUseBOM(true);
  				$objWriter->setSheetIndex(0);
  				$objWriter->save('php://output');
  				break;
  				case 'xls':
					// Redirect output to a client’s web browser (Excel5)
  				header('Content-Type: application/vnd.ms-excel');
  				header('Content-Disposition: attachment;filename="export.xls"');
  				header('Cache-Control: max-age=0');
  				$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
  				$objWriter->save('php://output');
  				break;
  				case 'xlsx':
					// Redirect output to a client’s web browser (Excel2007)
  				header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  				header('Content-Disposition: attachment;filename="export.xlsx"');
  				header('Cache-Control: max-age=0');
  				$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
  				$objWriter->save('php://output');
  				break;
  			}
  			exit;
  		}
  	} 
  	?>
  	<div class="wrap">
  		<h2><?php _e( "PHPExcel Export" ); ?></h2>
  		<?php 
  		global $_GET;
  		$author = $_GET[author];
  		$start_date = $_GET[start_date];
  		$end_date = $_GET[end_date];
  		?>
  		<form method='get' action="admin.php?page=spee-dashboard">
  			<input type="hidden" name='page' value="spee-dashboard"/>
  			<input type="hidden" name='noheader' value="1"/>
  			<div class="col-md-12">
  				<label>Pilih Author :</label>
  				<select name="author">
  					<?php 
  					$args = array(
  						'orderby'       => 'id', 
  						'order'         => 'ASC', 
  						'number'        => null,
  						'optioncount'   => false, 
  						'exclude_admin' => true, 
  						'show_fullname' => false,
  						'hide_empty'    => true,
  						'echo'          => false,
  						'style'         => 'none',
  						'html'          => false ); 

  					$author_list =  wp_list_authors( $args );
  					$author = explode(", ",$author_list);

  					foreach ($author as $value) { ?><option value="<?php echo $value; ?>"><?php echo $value; ?></option><?php } ?>
  				</select>
  			</div>
  			<div class="col-md-12">
  				<label>Pilih Tanggal Mulai :</label>
  				<input type="date" name='start_date' value="<?php echo $start_date ?>"/>
  			</div>
  			<div class="col-md-12">
  				<label>Pilih Tanggal Sampai :</label>
  				<input type="date" name='end_date' value="<?php echo $end_date ?>"/>
  			</div>
  			<input style="display:none" type="radio" name='format' id="formatCSV" value="csv" checked="checked"/>
  			<input type="submit" name='export' id="csvExport" value="Export"/>
  		</form>
  		<div class="footer-credit alignright">
  			<p>Thanks to awesome people at <a title="anang pratika" href="https://phpexcel.codeplex.com<" target="_blank" >PHPExcel</a>.</p>
  		</div>
  	</div>
  	<?php
  }
