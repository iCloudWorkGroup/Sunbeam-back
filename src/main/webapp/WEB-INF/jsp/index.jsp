<%@ page language="java" contentType="text/html; charset=utf-8"
	pageEncoding="utf-8"%>
<!doctype html>
<html>
	<head>
		<meta charset="utf-8">
		<title>Spreadsheet Online Office</title>
		<link rel="shortcut icon" type="image/x-icon" href="/${frontName}/favicon.ico"/>
		<link rel="stylesheet" type="text/css" href="/${frontName}/css/main.css"/>
		<link rel="stylesheet" type="text/css" href="/${frontName}/css/toolbar.css"/>
		<link rel="stylesheet" type="text/css" href="/${frontName}/css/widget.css"/>
	</head>
	<body>
		<input type="hidden" id="build" value=${build}>
		<input type="hidden" id ="excelId" value=${excelId}> 
		<!-- <button id="btn">test</button> -->
		<div class="topBar">
			<div class="file-control">
				<span>文件</span>
			</div>
			<div class="fui-control">
				<ul class="fui-control-list">
					<li><span id="homeTool" class="active">开始</span></li>
					<li><span id="reviewTool">审阅</span></li>
					<li><span><a id="download" href="<%=request.getContextPath() %>/download/${excelId}">下载</a></span></li>
					<!--  <li><span><a href="/${frontName}/excel.htm?m=save&excelId=${excelId}">保存</a></span></li> -->
					<li><span><a href="<%=request.getContextPath() %>/reopen/${excelId}">重新打开</a></span></li>

				</ul>
			</div>
			<div class="excel-name">
				<div class="textarea">Book Name</div>
			</div>
			<!--<div class="version">Version frontend : 0.6.0</div>
			<div class="version">Version java : 0.10.1</div>-->
		</div>
		<div class="toolBar" id="toolBar">
			<ul class="tabContainer homeToolContainer">
				<li class="fui-group" id="undoredoContainer">
					<span class="fui-container">
						<div class="fui-body">
							<span class="fui-layout">
								<div class="fui-transverse">
									<span class="ico-section" data-toolbar="redo">
										<div class="fui-cf-bg-extend-ico ico-redo"></div>
									</span>
								</div>
								<div class="fui-transverse">
									<span class="ico-section" data-toolbar="undo">
										<div class="fui-cf-bg-extend-ico ico-undo"></div>
									</span>
								</div>
							</span>
						</div>
						<div class="fui-title">撤销</div>
					</span>
					<span class="fui-separator"></span>
				</li>
				<li class="fui-group" id="shearPlateContainer">
					<span class="fui-container">
						<div class="fui-body">
							<span class="fui-layout">
								<div class="fui-section fui-alone" data-toolbar="paste">
									<div class="fui-cf-ico ico-paste"></div><div class="fui-cf-desc"><div class="fui-cf-text">粘贴</div></div>
								</div>
							</span>
							<span class="fui-layout">
								<div class="fui-section fui-transverse"  data-toolbar="cut">
									<div class="fui-transverse-model"><span class="fui-cf-ico ico-section-ico ico-shear"></span><span class="fui-cf-text">剪切</span></div>
								</div>
								<div class="fui-section fui-transverse" data-toolbar="copy">
									<div class="fui-transverse-model"><span class="fui-cf-ico ico-section-ico ico-copy"></span><span class="fui-cf-text">复制</span></div>
								</div>
							</span>
						</div>
						<div class="fui-title">剪贴板</div>
					</span>
					<span class="fui-separator"></span>
				</li>
				<li class="fui-group" id="contentFontContainer">
					<span class="fui-container">
						<div class="fui-body">
							<span class="fui-layout">
								<div class="fui-transverse">
									<span class="section" data-toolbar="fontfamily">
										<span class="fui-transverse-model fui-cf-fontfamily" id="fontShow">宋体</span>
										<span class="fui-transverse-model fui-cf-fontfamily-extend active">
											<span class="caret"></span>
										</span>
									</span>
									<span class="section"  data-toolbar="fontsize">
										<span class="fui-transverse-model fui-cf-fontsize" id="fontSizeShow">11</span>
										<span class="fui-transverse-model fui-cf-fontsize-extend active">
											<span class="caret"></span>
										</span>
									</span>
								</div>
								<div class="fui-transverse">
									<span class="ico-section" data-toolbar="bold" title="加粗"><span class="fui-cf-ico ico-weight"></span></span>
									<span class="ico-section" data-toolbar="italic"><span class="fui-cf-ico ico-italic"></span></span>
									<span class="ico-section" data-toolbar="border"><span class="fui-cf-ico ico-borderbottom ico-section-ico"></span><span class="ico-section-rightarrow"><span class="caret"></span></span></span>
									<span class="ico-section" data-toolbar="background"><span class="fui-cf-ico ico-fillbg ico-section-ico"></span><span class="ico-section-rightarrow"><span class="caret"></span></span></span>
									<span class="ico-section" data-toolbar="color"><span class="fui-cf-ico ico-fillcolor ico-section-ico"></span><span class="ico-section-rightarrow"><span class="caret"></span></span></span>
								</div>
							</span>
						</div>
						<div class="fui-title">字体</div>
					</span>
					<span class="fui-separator"></span>
				</li>
				<li class="fui-group" id="contentAlignContainer">
					<span class="fui-container">
						<div class="fui-body">
							<span class="fui-layout">
								<div class="fui-transverse">
									<span class="ico-section" data-align="top">
										<span class="fui-cf-ico ico-aligntop"></span>
									</span>
									<span class="ico-section" data-align="middle">
										<span class="fui-cf-ico ico-alignmiddle"></span>
									</span>
									<span class="ico-section" data-align="bottom">
										<span class="fui-cf-ico ico-alignbottom"></span>
									</span>
								</div>
								<div class="fui-transverse">
									<span class="ico-section" data-align="left">
										<span class="fui-cf-ico ico-alignleft"></span>
									</span>
									<span class="ico-section" data-align="center">
										<span class="fui-cf-ico ico-aligncenter"></span>
									</span>
									<span class="ico-section" data-align="right">
										<span class="fui-cf-ico ico-alignright"></span>
									</span>
								</div>
							</span>
						</div>
						<div class="fui-title">对齐方式</div>
					</span>
					<span class="fui-separator"></span>
				</li>
				<li class="fui-group">
					<span class="fui-container">
						<div class="fui-body" id="textFormatContainer">
							<span class="fui-layout">
								<div class="fui-section fui-alone" data-toolbar="format">
									<div class="fui-cf-ico ico-routine fui-cf-alone"></div>
									<div class="fui-cf-desc">
										<div class="fui-cf-text">文本格式</div>
										<div class="fui-cf-extend caret"></div>
									</div>
								</div>
							</span>
						</div>
						<div class="fui-title">数字</div>
					</span>
					<span class="fui-separator"></span>
				</li>
				<li class="fui-group">
					<span class="fui-container">
						<div class="fui-body" id="mergeCellContainer">
							<span class="fui-layout">
								<div class="fui-section fui-transverse" data-toolbar="merge" id="merge">
									<div class="fui-transverse-model">
										<span class="fui-cf-ico ico-section-ico ico-combincells"></span>
										<span class="fui-cf-text">合并单元格</span>
									</div>
								</div>
								<div class="fui-section fui-transverse" data-toolbar="split" id="split">
									<div class="fui-transverse-model">
										<span class="fui-cf-ico ico-section-ico ico-combincells"></span>
										<span class="fui-cf-text">拆分单元格</span>
									</div>
								</div>
							</span>
						</div>
						<div class="fui-title">单元格</div>
					</span>
					<span class="fui-separator"></span>
				</li>
				<li class="fui-group">
					<span class="fui-container">
						<div class="fui-body">
							<span class="fui-layout">
								<div class="fui-section fui-alone" data-toolbar="frozen">
									<div class="fui-cf-extend-ico ico-frozencustomized fui-cf-alone"></div>
									<div class="fui-cf-desc">
										<div class="fui-cf-text">冻结窗口</div>
										<div class="fui-cf-extend caret"></div>
									</div>
								</div>
							</span>
						</div>
						<div class="fui-title">视图</div>
					</span>
					<span class="fui-separator"></span>
				</li>
				<li class="fui-group">
					<span class="fui-container">
						<div class="fui-body" id="wordWrapContainer">
							<span class="fui-layout">
								<div class="fui-section fui-transverse" data-toolbar="wordwrap">
									<div class="fui-transverse-model">
										<span class="fui-cf-ico ico-section-ico ico-wordwrap"></span>
										<span class="fui-cf-text">自动换行</span>
									</div>
								</div>
							</span>
						</div>
						<div class="fui-title">文本对齐</div>
					</span>
					<span class="fui-separator"></span>
				</li>
				<li class="fui-group" id="reviewContainer">
					<span class="fui-container">
						<div class="fui-body">
							<span class="fui-layout">
								<div class="fui-section fui-alone" data-toolbar="addComment">
									<div class="fui-cf-comment-ico ico-commentadd fui-cf-alone"></div>
									<div class="fui-cf-desc">
										<div class="fui-cf-text">增加批注</div>
									</div>
								</div>
							</span>
							<span class="fui-layout">
								<div class="fui-section fui-alone" data-toolbar="editComment">
									<div class="fui-cf-comment-ico ico-commentedit fui-cf-alone"></div>
									<div class="fui-cf-desc">
										<div class="fui-cf-text">编辑批注</div>
									</div>
								</div>
							</span>
							<span class="fui-layout">
								<div class="fui-section fui-alone" data-toolbar="deleteComment">
									<div class="fui-cf-comment-ico ico-commentremove fui-cf-alone"></div>
									<div class="fui-cf-desc">
										<div class="fui-cf-text">删除批注</div>
									</div>
								</div>
							</span>
							<span class="fui-separator"></span>
						</div>
						<div class="fui-title">批注</div>
					</span>
				<li/>
				<li class="fui-group">
					<span class="fui-container">
						<div class="fui-body">
							<span class="fui-layout">
								<div class="fui-section fui-alone" data-toolbar="insert">
									<div class="fui-cf-bg-extend2-ico ico-insert fui-cf-alone"></div>
									<div class="fui-cf-desc">
										<div class="fui-cf-text">　插入　</div>
										<div class="fui-cf-extend caret"></div>
									</div>
								</div>
							</span>
							<span class="fui-layout">
								<div class="fui-section fui-alone" data-toolbar="delete">
									<div class="fui-cf-bg-extend2-ico ico-delete fui-cf-alone"></div>
									<div class="fui-cf-desc">
										<div class="fui-cf-text">　删除　</div>
										<div class="fui-cf-extend caret"></div>
									</div>
								</div>
							</span>
						</div>
						<div class="fui-title">行列</div>
					</span>
					<span class="fui-separator"></span>
				</li>
				<li class="fui-group" id="hideContainer">
					<span class="fui-container">
						<div class="fui-body" >
							<span class="fui-layout">
								<div class="fui-section fui-transverse" data-toolbar="hide">
									<div class="fui-cf-extend-ico ico-frozencustomized fui-cf-alone"></div>
									<div class="fui-cf-desc">
										<div class="fui-cf-text">隐藏</div>
									</div>
								</div>
							</span>
							<span class="fui-layout">
								<div class="fui-section fui-transverse" data-toolbar="cancelHide">
									<div class="fui-cf-extend-ico ico-frozencustomized fui-cf-alone"></div>
									<div class="fui-cf-desc">
										<div class="fui-cf-text">取消隐藏</div>
									</div>
								</div>
							</span>
						</div>
						<div class="fui-title">隐藏</div>
					</span>
					<span class="fui-separator"></span>
				</li>
			</ul>
		</div>
		<div style="top:130px;bottom:0;position:absolute;left:0;right:0;">
			<div id="spreadSheet" style="position:relative;height:100%;"></div>
		</div>
		<div class="widget-list">
			<div class="widget" data-widget="border">
				<div class="widget-panel">
					<ul class="widget-menu" id="funcBorder">
						<li data-border="bottom">
							<span class="fui-cf-ico ico-borderbottom ico-section-ico widget-ico"></span>
							<span class="widget-content">
								<div>下边框</div>
							</span>
						</li>
						<li data-border="top">
							<span class="fui-cf-ico ico-bordertop ico-section-ico widget-ico"></span>
							<span class="widget-content">
								<div>上边框</div>
							</span>
						</li>
						<li data-border="left">
							<span class="fui-cf-ico ico-borderleft ico-section-ico widget-ico"></span>
							<span class="widget-content">
								<div>左边框</div>
							</span>
						</li>
						<li data-border="right">
							<span class="fui-cf-ico ico-borderright ico-section-ico widget-ico"></span>
							<span class="widget-content">
								<div>右边框</div>
							</span>
						</li>
						<li data-border="none">
							<span class="fui-cf-ico ico-bordernone ico-section-ico widget-ico"></span>
							<span class="widget-content">
								<div>无边框</div>
							</span>
						</li>
						<li data-border="all">
							<span class="fui-cf-ico ico-borderall ico-section-ico widget-ico"></span>
							<span class="widget-content">
								<div>所有边框</div>
							</span>
						</li>
						<li data-border="outer">
							<span class="fui-cf-ico ico-borderouter ico-section-ico widget-ico"></span>
							<span class="widget-content">
								<div>外侧边框</div>
							</span>
						</li>
					</ul>
				</div>
			</div>
			<div class="widget" data-widget="format" id="contentFormat">
				<div class="widget-panel">
					<ul class="widget-menu">
						<li data-format="normal">
							<span class="fui-cf-ico ico-routine  widget-ico"></span>
							<span class="widget-content">
								<div class="widget-pad">常规</div>
							</span>
						</li>
						<li data-format="text">
							<span class="fui-cf-ico ico-text  widget-ico"></span>
							<span class="widget-content">
								<div class="widget-pad">文本</div>
							</span>
						</li>
						<li data-format="number-0">
							<span class="fui-cf-ico ico-numeric widget-ico"></span>
							<span class="widget-content">
								<div class="widget-pad">数字 0</div>
							</span>
						</li>
						<li data-format="number-1">
							<span class="fui-cf-ico ico-numeric widget-ico"></span>
							<span class="widget-content">
								<div class="widget-pad">数字 0.0</div>
							</span>
						</li>
						<li data-format="number-2">
							<span class="fui-cf-ico ico-numeric widget-ico"></span>
							<span class="widget-content">
								<div class="widget-pad">数字 0.00</div>
							</span>
						</li>
						<li data-format="number-3">
							<span class="fui-cf-ico ico-numeric widget-ico"></span>
							<span class="widget-content">
								<div class="widget-pad">数字 0.000</div>
							</span>
						</li>
						<li data-format="number-4">
							<span class="fui-cf-ico ico-numeric widget-ico"></span>
							<span class="widget-content">
								<div class="widget-pad">数字 0.0000</div>
							</span>
						</li>
						<li data-format="date-1">
							<span class="fui-cf-ico ico-date widget-ico"></span>
							<span class="widget-content">
								<div class="widget-pad">日期 1999/01/01</div>
							</span>
						</li>
						<li data-format="date-2">
							<span class="fui-cf-ico ico-date widget-ico"></span>
							<span class="widget-content">
								<div class="widget-pad">日期 1999年01月01日</div>
							</span>
						</li>
						<li data-format="date-3">
							<span class="fui-cf-ico ico-date widget-ico"></span>
							<span class="widget-content">
								<div class="widget-pad">日期 1999年01月</div>
							</span>
						</li>
						<li data-format="percent">
							<span class="fui-cf-ico ico-percent widget-ico"></span>
							<span class="widget-content">
								<div class="widget-pad">百分比</div>
							</span>
						</li>
						<li data-format="coin-1">
							<span class="fui-cf-ico ico-currency widget-ico"></span>
							<span class="widget-content">
								<div class="widget-pad">货币 $</div>
							</span>
						</li>
						<li data-format="coin-2">
							<span class="fui-cf-ico ico-currency widget-ico"></span>
							<span class="widget-content">
								<div class="widget-pad">货币 ¥</div>
							</span>
						</li>
					</ul>
				</div>
			</div>
			<div class="widget" data-widget="fontfamily" >
				<div class="widget-panel" style="max-height:200px">
					<ul class="font-list" style="min-width:160px" id="font">
						<li data-family="microsoft Yahei"><span style="font-family:microsoft Yahei">微软雅黑</span></li>
						<li data-family="SimSun"><span style="font-family:SimSun">宋体</span></li>
						<li data-family="SimHei"><span style="font-family:SimHei">黑体</span></li>
						<li data-family="Tahoma"><span style="font-family:Tahoma">Tahoma</span></li>
						<li data-family="Arial"><span style="font-family:Arial">Arial</span></li>
						<li data-family="楷体"><span style="font-family:楷体">楷体</span></li>
					</ul>
				</div>
			</div>
			<div class="widget" data-widget="fontsize">
				<div class="widget-panel" style="max-height:200px;">
					<ul class="font-list" id="fontSize">
						<li data-size="8"><span style="font-size:8pt">8</span></li>
						<li data-size="9"><span style="font-size:9pt">9</span></li>
						<li data-size="10"><span style="font-size:10pt">10</span></li>
						<li data-size="11"><span style="font-size:11pt">11</span></li>
						<li data-size="12"><span style="font-size:12pt">12</span></li>
						<li data-size="14"><span style="font-size:14pt">14</span></li>
						<li data-size="16"><span style="font-size:16pt">16</span></li>
						<li data-size="18"><span style="font-size:18pt">18</span></li>
						<li data-size="20"><span style="font-size:20pt">20</span></li>
						<li data-size="24"><span style="font-size:24pt">24</span></li>
						<li data-size="26"><span style="font-size:26pt">26</span></li>
						<li data-size="28"><span style="font-size:28pt">28</span></li>
						<li data-size="36"><span style="font-size:36pt">36</span></li>
					</ul>
				</div>
			</div>
			<div class="widget" data-widget="frozen" id="frozen">
				<div class="widget-panel">
					<ul class="widget-menu frozenBox" style="min-width:220px">
						<li data-frozen="unfrozen">
							<span class="fui-cf-extend-ico ico-frozencustomized widget-ico"></span>
							<span class="widget-content">
								<div class="widget-weight">取消冻结窗口</div>
								<div class="widget-desc">解除所有行和列锁定，以滚动整个工作表。</div>
							</span>
						</li>
						<li data-frozen="custom">
							<span class="fui-cf-extend-ico ico-frozencustomized widget-ico"></span>
							<span class="widget-content">
								<div class="widget-weight">冻结拆分窗口</div>
								<div class="widget-desc">滚动工作表其余部分时，保持行和列可见(基于当前的选择)</div>
							</span>
						</li>
						<li data-frozen="row">
							<span class="fui-cf-extend-ico ico-frozencustomized widget-ico"></span>
							<span class="widget-content">
								<div class="widget-weight">冻结首行</div>
								<div class="widget-desc">滚动工作表其余部分时，保持首行可见</div>
							</span>
						</li>
						<li data-frozen="col">
							<span class="fui-cf-extend-ico ico-frozencustomized widget-ico"></span>
							<span class="widget-content">
								<div class="widget-weight">冻结首列</div>
								<div class="widget-desc">滚动工作表其余部分时，保持行列可见</div>
							</span>
						</li>
					</ul>
				</div>
			</div>
			<div class="widget" data-widget="insert" id="insert">
				<div class="widget-panel">
					<ul class="widget-menu frozenBox" style="min-width:220px">
						<li data-type="row">
							<span class=" widget-ico"></span>
							<span class="widget-content">
								<div class="widget-weight">插入工作表行</div>
							</span>
						</li>
						<li data-type="column">
							<span class="widget-ico"></span>
							<span class="widget-content">
								<div class="widget-weight">插入工作表列</div>
							</span>
						</li>
					</ul>
				</div>
			</div>
			<div class="widget" data-widget="delete" id="delete">
				<div class="widget-panel">
					<ul class="widget-menu frozenBox" style="min-width:220px">
						<li data-type="row">
							<span class="widget-ico"></span>
							<span class="widget-content">
								<div class="widget-weight">删除工作表行</div>
							</span>
						</li>
						<li data-type="column">
							<span class="widget-ico"></span>
							<span class="widget-content">
								<div class="widget-weight">删除工作表列</div>
							</span>
						</li>
					</ul>
				</div>
			</div>
			<div class="widget" data-widget="background" id="fillColor">
				<div class="widget-panel">
					<div style="padding:3px;cursor:default;font-size:9pt">
						<div style="padding:2px;height:15px">
							<div style="float:left;">
								<a title="白色" class="color-box">
									<div class="color-body" style="background-color: rgb(0, 0, 0);"></div>
								</a>
							</div>
							<div style="margin-left:20px;">自动</div>
						</div>
						<div style="padding:2px 0;background:#fff;border-top:1px solid #e1e1e1;border-bottom:1px solid #e1e1e1">
							<table class="" cellspacing="0" cellpadding="0">
								<tbody>
									<tr>
										<td arrayposition="0">
											<a title="白色" class="color-box">
												<div class="color-body" style="background-color: rgb(255, 255, 255);"></div>
											</a>
										</td>
										<td arrayposition="1">
											<a title="黑色" class="color-box">
												<div class="color-body" style="background-color: rgb(0, 0, 0);"></div>
											</a>
										</td>
										<td arrayposition="2">
											<a title="茶色" class="color-box">
												<div class="color-body" style="background-color: rgb(238, 236, 225);"></div>
											</a>
										</td>
										<td arrayposition="3">
											<a title="深蓝色" class="color-box">
												<div class="color-body" style="background-color: rgb(31, 73, 125);"></div>
											</a>
										</td>
										<td arrayposition="4">
											<a title="蓝色" class="color-box">
												<div class="color-body" style="background-color: rgb(79, 129, 189);"></div>
											</a>
										</td>
										<td arrayposition="5">
											<a title="红色" class="color-box">
												<div class="color-body" style="background-color: rgb(192, 80, 77);"></div>
											</a>
										</td>
										<td arrayposition="6">
											<a title="橄榄色" class="color-box">
												<div class="color-body" style="background-color: rgb(155, 187, 89);"></div>
											</a>
										</td>
										<td arrayposition="7">
											<a title="紫色" class="color-box">
												<div class="color-body" style="background-color: rgb(128, 100, 162);"></div>
											</a>
										</td>
										<td arrayposition="8">
											<a title="水绿色" class="color-box">
												<div class="color-body" style="background-color: rgb(75, 172, 198);"></div>
											</a>
										</td>
										<td arrayposition="9">
											<a title="橙色" class="color-box">
												<div class="color-body" style="background-color: rgb(247, 150, 70);"></div>
											</a>
										</td>
									</tr>
									<tr>
										<td arrayposition="10">
											<a title="深白色 5%" class="color-box">
												<div class="color-body" style="background-color: rgb(242, 242, 242);"></div>
											</a>
										</td>
										<td arrayposition="11">
											<a title="浅黑色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(127, 127, 127);"></div>
											</a>
										</td>
										<td arrayposition="12">
											<a title="深茶色 10%" class="color-box">
												<div class="color-body" style="background-color: rgb(221, 217, 195);">
													
												</div>
											</a>
										</td>
										<td arrayposition="13">
											<a title="较浅的深蓝色 80%" class="color-box">
												<div class="color-body" style="background-color: rgb(198, 217, 240);"></div>
											</a>
										</td>
										<td arrayposition="14">
											<a title="浅蓝色 80%" class="color-box">
												<div class="color-body" style="background-color: rgb(219, 229, 241);"></div>
											</a>
										</td>
										<td arrayposition="15">
											<a title="浅红色 80%" class="color-box">
												<div class="color-body" style="background-color: rgb(242, 220, 219);"></div>
											</a>
										</td>
										<td arrayposition="16">
											<a title="较浅的橄榄色 80%" class="color-box">
												<div class="color-body" style="background-color: rgb(235, 241, 221);"></div>
											</a>
										</td>
										<td arrayposition="17">
											<a title="浅紫色 80%" class="color-box">
												<div class="color-body" style="background-color: rgb(229, 224, 236);"></div>
											</a>
										</td>
										<td arrayposition="18">
											<a title="浅绿色 80%" class="color-box">
												<div class="color-body" style="background-color: rgb(219, 238, 243);">
													
												</div>
											</a>
										</td>
										<td arrayposition="19">
											<a title="浅橙色 80%" class="color-box"><div class="color-body" style="background-color: rgb(253, 234, 218);"></div></a>
										</td>
									</tr>
									<tr>
										<td arrayposition="20">
											<a title="深白色 15%" class="color-box">
												<div class="color-body" style="background-color: rgb(216, 216, 216);">
													
												</div>
											</a>
										</td>
										<td arrayposition="21">
											<a title="浅黑色 35%" class="color-box">
												<div class="color-body" style="background-color: rgb(89, 89, 89);">
													
												</div>
											</a>
										</td>
										<td arrayposition="22">
											<a title="深茶色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(196, 189, 151);">
													
												</div>
											</a>
										</td>
										<td arrayposition="23">
											<a title="较浅的深蓝色 60%" class="color-box">
												<div class="color-body" style="background-color: rgb(141, 179, 226);">
													
												</div>
											</a>
										</td>
										<td arrayposition="24">
											<a title="浅蓝色 60%" class="color-box">
												<div class="color-body" style="background-color: rgb(184, 204, 228);">
													
												</div>
											</a>
										</td>
										<td arrayposition="25">
											<a title="浅红色 60%" class="color-box">
												<div class="color-body" style="background-color: rgb(229, 185, 183);">
													
												</div>
											</a>
										</td>
										<td arrayposition="26">
											<a title="较浅的橄榄色 60%" class="color-box">
												<div class="color-body" style="background-color: rgb(215, 227, 188);">
													
												</div>
											</a>
										</td>
										<td arrayposition="27">
											<a title="浅紫色 60%" class="color-box">
												<div class="color-body" style="background-color: rgb(204, 193, 217);">
													
												</div>
											</a>
										</td>
										<td arrayposition="28">
											<a title="浅绿色 80%" class="color-box">
												<div class="color-body" style="background-color: rgb(183, 221, 232);">
													
												</div>
											</a>
										</td>
										<td arrayposition="29">
											<a title="浅橙色 60%" class="color-box">
												<div class="color-body" style="background-color: rgb(251, 213, 181);">
													
												</div>
											</a>
										</td>
									</tr>
									<tr>
										<td arrayposition="30">
											<a title="深白色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(191, 191, 191);">
													
												</div>
											</a>
										</td>
										<td arrayposition="31">
											<a title="浅黑色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(63, 63, 63);">
													
												</div>
											</a>
										</td>
										<td arrayposition="32">
											<a title="深茶色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(147, 137, 83);">
													
												</div>
											</a>
										</td>
										<td arrayposition="33">
											<a title="较浅的深蓝色 40%" class="color-box">
												<div class="color-body" style="background-color: rgb(84, 141, 212);">
													
												</div>
											</a>
										</td>
										<td arrayposition="34">
											<a title="浅蓝色 40%" class="color-box">
												<div class="color-body" style="background-color: rgb(149, 179, 215);">
													
												</div>
											</a>
										</td>
										<td arrayposition="35">
											<a title="浅红色 40%" class="color-box">
												<div class="color-body" style="background-color: rgb(217, 150, 148);">
													
												</div>
											</a>
										</td>
										<td arrayposition="36">
											<a title="较浅的橄榄色 40%" class="color-box">
												<div class="color-body" style="background-color: rgb(195, 214, 155);">
													
												</div>
											</a>
										</td>
										<td arrayposition="37">
											<a title="浅紫色 40%" class="color-box">
												<div class="color-body" style="background-color: rgb(178, 162, 199);">
													
												</div>
											</a>
										</td>
										<td arrayposition="38">
											<a title="浅绿色 40%" class="color-box">
												<div class="color-body" style="background-color: rgb(146, 205, 220);">
													
												</div>
											</a>
										</td>
										<td arrayposition="39">
											<a title="浅橙色 40%" class="color-box">
												<div class="color-body" style="background-color: rgb(250, 192, 143);">
													
												</div>
											</a>
										</td>
									</tr>
									<tr>
										<td arrayposition="40">
											<a title="深白色 5%" class="color-box">
												<div class="color-body" style="background-color: rgb(165, 165, 165);">
													
												</div>
											</a>
										</td>
										<td arrayposition="41">
											<a title="浅黑色 15%" class="color-box">
												<div class="color-body" style="background-color: rgb(38, 38, 38);">
													
												</div>
											</a>
										</td>
										<td arrayposition="42">
											<a title="深茶色 75%" class="color-box">
												<div class="color-body" style="background-color: rgb(73, 68, 41);">
													
												</div>
											</a>
										</td>
										<td arrayposition="43">
											<a title="较浅的深蓝色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(23, 54, 93);">
													
												</div>
											</a>
										</td>
										<td arrayposition="44">
											<a title="深黑色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(54, 96, 146);">
													
												</div>
											</a>
										</td>
										<td arrayposition="45">
											<a title="深红色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(149, 55, 52);">
													
												</div>
											</a>
										</td>
										<td arrayposition="46">
											<a title="较深的橄榄色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(118, 146, 60);">
													
												</div>
											</a>
										</td>
										<td arrayposition="47">
											<a title="深紫色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(95, 73, 122);">
													
												</div>
											</a>
										</td>
										<td arrayposition="48">
											<a title="暗绿色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(49, 132, 155);">
													
												</div>
											</a>
										</td>
										<td arrayposition="49">
											<a title="深橙色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(227, 108, 9);">
													
												</div>
											</a>
										</td>
									</tr>
									<tr>
										<td arrayposition="50">
											<a title="深白色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(127, 127, 127);">
													
												</div>
											</a>
										</td>
										<td arrayposition="51">
											<a title="浅黑色 5%" class="color-box">
												<div class="color-body" style="background-color: rgb(12, 12, 12);">
													
												</div>
											</a>
										</td>
										<td arrayposition="52">
											<a title="深茶色 90%" class="color-box">
												<div class="color-body" style="background-color: rgb(29, 27, 16);">
													
												</div>
											</a>
										</td>
										<td arrayposition="53">
											<a title="较浅的深蓝色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(15, 36, 62);">
													
												</div>
											</a>
										</td>
										<td arrayposition="54">
											<a title="深黑色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(36, 64, 97);">
													
												</div>
											</a>
										</td>
										<td arrayposition="55">
											<a title="深红色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(99, 36, 35);">
													
												</div>
											</a>
										</td>
										<td arrayposition="56">
											<a title="较深的橄榄色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(79, 97, 40);">
													
												</div>
											</a>
										</td>
										<td arrayposition="57">
											<a title="深紫色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(63, 49, 81);">
													
												</div>
											</a>
										</td>
										<td arrayposition="58">
											<a title="暗绿色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(32, 88, 103);">
													
												</div>
											</a>
										</td>
										<td arrayposition="59">
											<a title="深橙色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(152, 72, 6);">
													
												</div>
											</a>
										</td>
									</tr>
								</tbody>
							</table>
						</div>
						<div>
							<div style="padding:6px;background:#e1e1e1;">标准颜色</div>
							<div style="padding-top:2px">
								<table cellspacing="0" cellpadding="0">
									<tbody>
										<tr>
											<td arrayposition="0">
												<a title="深红色" class="color-box">
													<div class="color-body" style="background-color: rgb(192, 0, 0);">
														
													</div>
												</a>
											</td>
											<td arrayposition="1">
												<a title="红色" class="color-box">
													<div class="color-body" style="background-color: rgb(255, 0, 0);">
														
													</div>
												</a>
											</td>
											<td arrayposition="2">
												<a title="橙色" class="color-box">
													<div class="color-body" style="background-color: rgb(255, 192, 0);">
														
													</div>
												</a>
											</td>
											<td arrayposition="3">
												<a title="黄色" class="color-box">
													<div class="color-body" style="background-color: rgb(255, 255, 0);">
														
													</div>
												</a>
											</td>
											<td arrayposition="4">
												<a title="浅绿色" class="color-box">
													<div class="color-body" style="background-color: rgb(146, 208, 80);">
														
													</div>
												</a>
											</td>
											<td arrayposition="5">
												<a title="深绿色" class="color-box">
													<div class="color-body" style="background-color: rgb(0, 176, 80);">
														
													</div>
												</a>
											</td>
											<td arrayposition="6">
												<a title="浅蓝色" class="color-box">
													<div class="color-body" style="background-color: rgb(0, 176, 240);">
														
													</div>
												</a>
											</td>
											<td arrayposition="7">
												<a title="蓝色" class="color-box">
													<div class="color-body" style="background-color: rgb(0, 112, 192);">
														
													</div>
												</a>
											</td>
											<td arrayposition="8">
												<a title="深蓝色" class="color-box">
													<div class="color-body" style="background-color: rgb(0, 32, 96);">
														
													</div>
												</a>
											</td>
											<td arrayposition="9">
												<a title="紫色" class="color-box">
													<div class="color-body" style="background-color: rgb(112, 48, 160);">
														
													</div>
												</a>
											</td>
										</tr>
									</tbody>
								</table>
							</div>
						</div>
					</div>
				</div>
			</div>
			<div class="widget" data-widget="color" id="fontColor">
				<div class="widget-panel">
					<div style="padding:3px;cursor:default;font-size:9pt">
						<div style="padding:2px;height:15px">
							<div style="float:left;">
								<a title="白色" class="color-box">
									<div class="color-body" style="background-color: rgb(0, 0, 0);"></div>
								</a>
							</div>
							<div style="margin-left:20px;">自动</div>
						</div>
						<div style="padding:2px 0;background:#fff;border-top:1px solid #e1e1e1;border-bottom:1px solid #e1e1e1">
							<table class="" cellspacing="0" cellpadding="0">
								<tbody>
									<tr>
										<td arrayposition="0">
											<a title="白色" class="color-box">
												<div class="color-body" style="background-color: rgb(255, 255, 255);"></div>
											</a>
										</td>
										<td arrayposition="1">
											<a title="黑色" class="color-box">
												<div class="color-body" style="background-color: rgb(0, 0, 0);"></div>
											</a>
										</td>
										<td arrayposition="2">
											<a title="茶色" class="color-box">
												<div class="color-body" style="background-color: rgb(238, 236, 225);"></div>
											</a>
										</td>
										<td arrayposition="3">
											<a title="深蓝色" class="color-box">
												<div class="color-body" style="background-color: rgb(31, 73, 125);"></div>
											</a>
										</td>
										<td arrayposition="4">
											<a title="蓝色" class="color-box">
												<div class="color-body" style="background-color: rgb(79, 129, 189);"></div>
											</a>
										</td>
										<td arrayposition="5">
											<a title="红色" class="color-box">
												<div class="color-body" style="background-color: rgb(192, 80, 77);"></div>
											</a>
										</td>
										<td arrayposition="6">
											<a title="橄榄色" class="color-box">
												<div class="color-body" style="background-color: rgb(155, 187, 89);"></div>
											</a>
										</td>
										<td arrayposition="7">
											<a title="紫色" class="color-box">
												<div class="color-body" style="background-color: rgb(128, 100, 162);"></div>
											</a>
										</td>
										<td arrayposition="8">
											<a title="水绿色" class="color-box">
												<div class="color-body" style="background-color: rgb(75, 172, 198);"></div>
											</a>
										</td>
										<td arrayposition="9">
											<a title="橙色" class="color-box">
												<div class="color-body" style="background-color: rgb(247, 150, 70);"></div>
											</a>
										</td>
									</tr>
									<tr>
										<td arrayposition="10">
											<a title="深白色 5%" class="color-box">
												<div class="color-body" style="background-color: rgb(242, 242, 242);"></div>
											</a>
										</td>
										<td arrayposition="11">
											<a title="浅黑色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(127, 127, 127);"></div>
											</a>
										</td>
										<td arrayposition="12">
											<a title="深茶色 10%" class="color-box">
												<div class="color-body" style="background-color: rgb(221, 217, 195);">
													
												</div>
											</a>
										</td>
										<td arrayposition="13">
											<a title="较浅的深蓝色 80%" class="color-box">
												<div class="color-body" style="background-color: rgb(198, 217, 240);"></div>
											</a>
										</td>
										<td arrayposition="14">
											<a title="浅蓝色 80%" class="color-box">
												<div class="color-body" style="background-color: rgb(219, 229, 241);"></div>
											</a>
										</td>
										<td arrayposition="15">
											<a title="浅红色 80%" class="color-box">
												<div class="color-body" style="background-color: rgb(242, 220, 219);"></div>
											</a>
										</td>
										<td arrayposition="16">
											<a title="较浅的橄榄色 80%" class="color-box">
												<div class="color-body" style="background-color: rgb(235, 241, 221);"></div>
											</a>
										</td>
										<td arrayposition="17">
											<a title="浅紫色 80%" class="color-box">
												<div class="color-body" style="background-color: rgb(229, 224, 236);"></div>
											</a>
										</td>
										<td arrayposition="18">
											<a title="浅绿色 80%" class="color-box">
												<div class="color-body" style="background-color: rgb(219, 238, 243);">
													
												</div>
											</a>
										</td>
										<td arrayposition="19">
											<a title="浅橙色 80%" class="color-box">
												<div class="color-body" style="background-color: rgb(253, 234, 218);">
													
												</div>
											</a>
										</td>
									</tr>
									<tr>
										<td arrayposition="20">
											<a title="深白色 15%" class="color-box">
												<div class="color-body" style="background-color: rgb(216, 216, 216);">
													
												</div>
											</a>
										</td>
										<td arrayposition="21">
											<a title="浅黑色 35%" class="color-box">
												<div class="color-body" style="background-color: rgb(89, 89, 89);">
													
												</div>
											</a>
										</td>
										<td arrayposition="22">
											<a title="深茶色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(196, 189, 151);">
													
												</div>
											</a>
										</td>
										<td arrayposition="23">
											<a title="较浅的深蓝色 60%" class="color-box">
												<div class="color-body" style="background-color: rgb(141, 179, 226);">
													
												</div>
											</a>
										</td>
										<td arrayposition="24">
											<a title="浅蓝色 60%" class="color-box">
												<div class="color-body" style="background-color: rgb(184, 204, 228);">
													
												</div>
											</a>
										</td>
										<td arrayposition="25">
											<a title="浅红色 60%" class="color-box">
												<div class="color-body" style="background-color: rgb(229, 185, 183);">
													
												</div>
											</a>
										</td>
										<td arrayposition="26">
											<a title="较浅的橄榄色 60%" class="color-box">
												<div class="color-body" style="background-color: rgb(215, 227, 188);">
													
												</div>
											</a>
										</td>
										<td arrayposition="27">
											<a title="浅紫色 60%" class="color-box">
												<div class="color-body" style="background-color: rgb(204, 193, 217);">
													
												</div>
											</a>
										</td>
										<td arrayposition="28">
											<a title="浅绿色 80%" class="color-box">
												<div class="color-body" style="background-color: rgb(183, 221, 232);">
													
												</div>
											</a>
										</td>
										<td arrayposition="29">
											<a title="浅橙色 60%" class="color-box">
												<div class="color-body" style="background-color: rgb(251, 213, 181);">
													
												</div>
											</a>
										</td>
									</tr>
									<tr>
										<td arrayposition="30">
											<a title="深白色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(191, 191, 191);">
													
												</div>
											</a>
										</td>
										<td arrayposition="31">
											<a title="浅黑色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(63, 63, 63);">
													
												</div>
											</a>
										</td>
										<td arrayposition="32">
											<a title="深茶色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(147, 137, 83);">
													
												</div>
											</a>
										</td>
										<td arrayposition="33">
											<a title="较浅的深蓝色 40%" class="color-box">
												<div class="color-body" style="background-color: rgb(84, 141, 212);">
													
												</div>
											</a>
										</td>
										<td arrayposition="34">
											<a title="浅蓝色 40%" class="color-box">
												<div class="color-body" style="background-color: rgb(149, 179, 215);">
													
												</div>
											</a>
										</td>
										<td arrayposition="35">
											<a title="浅红色 40%" class="color-box">
												<div class="color-body" style="background-color: rgb(217, 150, 148);">
													
												</div>
											</a>
										</td>
										<td arrayposition="36">
											<a title="较浅的橄榄色 40%" class="color-box">
												<div class="color-body" style="background-color: rgb(195, 214, 155);">
													
												</div>
											</a>
										</td>
										<td arrayposition="37">
											<a title="浅紫色 40%" class="color-box">
												<div class="color-body" style="background-color: rgb(178, 162, 199);">
													
												</div>
											</a>
										</td>
										<td arrayposition="38">
											<a title="浅绿色 40%" class="color-box">
												<div class="color-body" style="background-color: rgb(146, 205, 220);">
													
												</div>
											</a>
										</td>
										<td arrayposition="39">
											<a title="浅橙色 40%" class="color-box">
												<div class="color-body" style="background-color: rgb(250, 192, 143);">
													
												</div>
											</a>
										</td>
									</tr>
									<tr>
										<td arrayposition="40">
											<a title="深白色 5%" class="color-box">
												<div class="color-body" style="background-color: rgb(165, 165, 165);">
													
												</div>
											</a>
										</td>
										<td arrayposition="41">
											<a title="浅黑色 15%" class="color-box">
												<div class="color-body" style="background-color: rgb(38, 38, 38);">
													
												</div>
											</a>
										</td>
										<td arrayposition="42">
											<a title="深茶色 75%" class="color-box">
												<div class="color-body" style="background-color: rgb(73, 68, 41);">
													
												</div>
											</a>
										</td>
										<td arrayposition="43">
											<a title="较浅的深蓝色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(23, 54, 93);">
													
												</div>
											</a>
										</td>
										<td arrayposition="44">
											<a title="深黑色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(54, 96, 146);">
													
												</div>
											</a>
										</td>
										<td arrayposition="45">
											<a title="深红色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(149, 55, 52);">
													
												</div>
											</a>
										</td>
										<td arrayposition="46">
											<a title="较深的橄榄色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(118, 146, 60);">
													
												</div>
											</a>
										</td>
										<td arrayposition="47">
											<a title="深紫色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(95, 73, 122);">
													
												</div>
											</a>
										</td>
										<td arrayposition="48">
											<a title="暗绿色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(49, 132, 155);">
													
												</div>
											</a>
										</td>
										<td arrayposition="49">
											<a title="深橙色 25%" class="color-box">
												<div class="color-body" style="background-color: rgb(227, 108, 9);">
													
												</div>
											</a>
										</td>
									</tr>
									<tr>
										<td arrayposition="50">
											<a title="深白色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(127, 127, 127);">
													
												</div>
											</a>
										</td>
										<td arrayposition="51">
											<a title="浅黑色 5%" class="color-box">
												<div class="color-body" style="background-color: rgb(12, 12, 12);">
													
												</div>
											</a>
										</td>
										<td arrayposition="52">
											<a title="深茶色 90%" class="color-box">
												<div class="color-body" style="background-color: rgb(29, 27, 16);">
													
												</div>
											</a>
										</td>
										<td arrayposition="53">
											<a title="较浅的深蓝色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(15, 36, 62);">
													
												</div>
											</a>
										</td>
										<td arrayposition="54">
											<a title="深黑色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(36, 64, 97);">
													
												</div>
											</a>
										</td>
										<td arrayposition="55">
											<a title="深红色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(99, 36, 35);">
													
												</div>
											</a>
										</td>
										<td arrayposition="56">
											<a title="较深的橄榄色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(79, 97, 40);">
													
												</div>
											</a>
										</td>
										<td arrayposition="57">
											<a title="深紫色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(63, 49, 81);">
													
												</div>
											</a>
										</td>
										<td arrayposition="58">
											<a title="暗绿色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(32, 88, 103);">
													
												</div>
											</a>
										</td>
										<td arrayposition="59">
											<a title="深橙色 50%" class="color-box">
												<div class="color-body" style="background-color: rgb(152, 72, 6);">
													
												</div>
											</a>
										</td>
									</tr>
								</tbody>
							</table>
						</div>
						<div>
							<div style="padding:6px;background:#e1e1e1;">标准颜色</div>
							<div style="padding-top:2px">
								<table cellspacing="0" cellpadding="0">
									<tbody>
										<tr>
											<td arrayposition="0">
												<a title="深红色" class="color-box">
													<div class="color-body" style="background-color: rgb(192, 0, 0);">
														
													</div>
												</a>
											</td>
											<td arrayposition="1">
												<a title="红色" class="color-box">
													<div class="color-body" style="background-color: rgb(255, 0, 0);">
														
													</div>
												</a>
											</td>
											<td arrayposition="2">
												<a title="橙色" class="color-box">
													<div class="color-body" style="background-color: rgb(255, 192, 0);">
														
													</div>
												</a>
											</td>
											<td arrayposition="3">
												<a title="黄色" class="color-box">
													<div class="color-body" style="background-color: rgb(255, 255, 0);">
														
													</div>
												</a>
											</td>
											<td arrayposition="4">
												<a title="浅绿色" class="color-box">
													<div class="color-body" style="background-color: rgb(146, 208, 80);">
														
													</div>
												</a>
											</td>
											<td arrayposition="5">
												<a title="深绿色" class="color-box">
													<div class="color-body" style="background-color: rgb(0, 176, 80);">
														
													</div>
												</a>
											</td>
											<td arrayposition="6">
												<a title="浅蓝色" class="color-box">
													<div class="color-body" style="background-color: rgb(0, 176, 240);">
														
													</div>
												</a>
											</td>
											<td arrayposition="7">
												<a title="蓝色" class="color-box">
													<div class="color-body" style="background-color: rgb(0, 112, 192);">
														
													</div>
												</a>
											</td>
											<td arrayposition="8">
												<a title="深蓝色" class="color-box">
													<div class="color-body" style="background-color: rgb(0, 32, 96);">
														
													</div>
												</a>
											</td>
											<td arrayposition="9">
												<a title="紫色" class="color-box">
													<div class="color-body" style="background-color: rgb(112, 48, 160);">
														
													</div>
												</a>
											</td>
										</tr>
									</tbody>
								</table>
							</div>
						</div>
					</div>
				</div>
			</div>
		</div>
		<script type="text/javascript" data-main='/${frontName}/app' src="/${frontName}/js/lib/require.js">
		</script>
	</body>
</html>
