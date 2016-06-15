CREATE TABLE `cost` (
	`order` varchar(50) NOT NULL COMMENT '订单',
	`month` int(10) NOT NULL COMMENT '年月',
	`price` varchar(50) NOT NULL COMMENT '单价',
	`num` int(10) NOT NULL COMMENT '订单数',
	`status` varchar(50) DEFAULT '0' COMMENT '0为未完款，1为完款',
	`stage` varchar(20) NOT NULL COMMENT '工序名字'
) ENGINE=InnoDB DEFAULT CHARSET=utf8
