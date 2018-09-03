<?php ?>

<!DOCTYPE html>
<html>
<body>
	<div><img src="<?php print $logo; ?>" class="logo-img"/>
	<div class="site-name"><?php print $site_name; ?></div></div>
<div class="header-reg">
<div class="breadcrum-reg"><?php print $breadcrumb; ?></div>
<?php print render($page['header']) ?>
</div>

	<div class="sub-header">
	<?php print render($page['sub-header'])?>
	</div>
<div class="center-reg">
	<div class="content-re">
	<?php print render($page['content'])?>
	</div>
	<div class="right-sidebar"><?php print render($page['right-sidebar']) ?></div>
</div>
<div class="footer-re">
<?php print render($page['foot'])?>
</div>
</body>
</html>