from PyQt5.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QMessageBox,
    QApplication
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QClipboard
from core.license import LicenseManager
import logging
import traceback

class RegisterDialog(QDialog):
    def __init__(self, license_manager, is_permanent=False):
        try:
            logging.info("初始化注册对话框...")
            super().__init__(None)  # 不设置父窗口，使其成为独立窗口
            
            # 确保窗口显示在最前面
            self.setWindowFlags(
                Qt.WindowCloseButtonHint | 
                Qt.WindowStaysOnTopHint |  # 保持窗口在最前
                Qt.Window  # 使其成为独立窗口
            )
            
            self.license_manager = license_manager
            self.is_permanent = is_permanent
            
            # 检查当前会员状态
            self.is_registered, self.is_vip, self.is_permanent_member, self.expire_time = self.license_manager.is_registered()
            
            self.init_ui()
            logging.info("注册对话框初始化完成")
            
        except Exception as e:
            logging.error(f"注册对话框初始化失败: {str(e)}")
            logging.error(traceback.format_exc())
            raise
        
    def init_ui(self):
        """初始化UI"""
        try:
            logging.info("开始初始化注册对话框UI...")
            self.setWindowTitle("会员激活" if not self.is_permanent else "永久会员激活")
            self.setFixedSize(400, 300)  # 增加高度以容纳VIP说明
            
            # 设置样式
            self.setStyleSheet("""
                QDialog {
                    background-color: white;
                }
                QLabel {
                    color: #333333;
                    font-size: 13px;
                }
                QLabel#vip_label {
                    color: #FFD700;
                    font-size: 14px;
                    font-weight: bold;
                }
                QLabel#price_label {
                    color: #FF4500;
                    font-size: 16px;
                    font-weight: bold;
                }
                QLabel#member_status {
                    color: #4CAF50;
                    font-size: 14px;
                    font-weight: bold;
                }
            """)
            
            layout = QVBoxLayout()
            layout.setSpacing(15)
            layout.setContentsMargins(20, 20, 20, 20)
            
            # 添加会员状态显示
            if self.is_vip:
                status_layout = QHBoxLayout()
                status_label = QLabel("当前状态：")
                member_status = QLabel("永久会员" if self.is_permanent_member else f"年费会员（到期时间：{self.expire_time}）")
                member_status.setObjectName("member_status")
                status_layout.addWidget(status_label)
                status_layout.addWidget(member_status)
                status_layout.addStretch()
                layout.addLayout(status_layout)
            
            # 添加会员类型标题
            title_label = QLabel("年费会员" if not self.is_permanent else "永久会员")
            title_label.setObjectName("vip_label")
            layout.addWidget(title_label)
            
            # 添加价格显示
            price_layout = QHBoxLayout()
            yearly_price = QLabel("年费会员 ￥88/年")
            permanent_price = QLabel("永久会员 ￥129")
            
            # 设置当前选择的价格为粗体
            if self.is_permanent:
                permanent_price.setStyleSheet("color: #FF4500; font-size: 16px; font-weight: bold;")
                yearly_price.setStyleSheet("color: #666666; font-size: 14px;")
            else:
                yearly_price.setStyleSheet("color: #FF4500; font-size: 16px; font-weight: bold;")
                permanent_price.setStyleSheet("color: #666666; font-size: 14px;")
            
            price_layout.addWidget(yearly_price)
            price_layout.addSpacing(20)  # 添加间距
            price_layout.addWidget(permanent_price)
            price_layout.addStretch()
            layout.addLayout(price_layout)
            
            # 添加功能说明
            features_label = QLabel(
                "• 无限制分析群聊数量\n"
                "• 显示所有重复成员\n"
                "• 支持导出分析结果"
            )
            layout.addWidget(features_label)
            
            # 用户ID显示
            user_id_layout = QHBoxLayout()
            user_id_label = QLabel("用户ID:")
            self.user_id_edit = QLineEdit()
            self.user_id_edit.setReadOnly(True)
            self.user_id_edit.setText(self.license_manager.get_user_id())
            copy_btn = QPushButton("复制用户ID")
            copy_btn.setObjectName("copy_btn")
            copy_btn.clicked.connect(self.copy_user_id)
            
            user_id_layout.addWidget(user_id_label)
            user_id_layout.addWidget(self.user_id_edit)
            user_id_layout.addWidget(copy_btn)
            
            # 激活码输入
            license_layout = QHBoxLayout()
            license_label = QLabel("激活码:")
            self.license_edit = QLineEdit()
            self.license_edit.setPlaceholderText("请输入激活码")
            
            license_layout.addWidget(license_label)
            license_layout.addWidget(self.license_edit)
            
            # 按钮
            button_layout = QHBoxLayout()
            register_btn = QPushButton("激活")
            cancel_btn = QPushButton("取消")
            
            register_btn.clicked.connect(self.register)
            cancel_btn.clicked.connect(self.reject)
            
            button_layout.addStretch()
            button_layout.addWidget(register_btn)
            button_layout.addWidget(cancel_btn)
            
            # 添加所有布局
            layout.addLayout(user_id_layout)
            layout.addLayout(license_layout)
            layout.addStretch()
            layout.addLayout(button_layout)
            
            self.setLayout(layout)
            logging.info("注册对话框UI初始化完成")
            
        except Exception as e:
            logging.error(f"注册对话框UI初始化失败: {str(e)}")
            logging.error(traceback.format_exc())
            raise
        
    def showEvent(self, event):
        """窗口显示事件"""
        try:
            logging.info("注册对话框开始显示...")
            super().showEvent(event)
            # 确保窗口显示在屏幕中央
            screen = QApplication.primaryScreen().geometry()
            x = (screen.width() - self.width()) // 2
            y = (screen.height() - self.height()) // 2
            self.move(x, y)
            
            # 确保窗口在最前面
            self.raise_()
            self.activateWindow()
            logging.info("注册对话框显示完成")
            
        except Exception as e:
            logging.error(f"注册对话框显示失败: {str(e)}")
            logging.error(traceback.format_exc())
        
    def copy_user_id(self):
        """复制用户ID到剪贴板"""
        try:
            logging.info("复制用户ID到剪贴板...")
            clipboard = QApplication.clipboard()
            clipboard.setText(self.user_id_edit.text())
            QMessageBox.information(self, "提示", "用户ID已复制到剪贴板！")
            logging.info("用户ID复制成功")
            
        except Exception as e:
            logging.error(f"复制用户ID失败: {str(e)}")
            logging.error(traceback.format_exc())
            QMessageBox.warning(self, "错误", "复制失败！")
        
    def register(self):
        """处理注册"""
        try:
            logging.info("开始处理注册...")
            license_code = self.license_edit.text().strip()
            if not license_code:
                logging.warning("用户未输入激活码")
                QMessageBox.warning(self, "错误", "请输入激活码！")
                return
            
            logging.info("验证激活码...")
            success, is_vip, is_permanent = self.license_manager.verify_license(license_code)
            if success:
                if is_permanent == self.is_permanent:  # 确认激活码类型匹配
                    logging.info(f"注册成功，会员类型：{'永久' if is_permanent else '年费'}")
                    message = "永久会员激活成功！" if is_permanent else "年费会员激活成功！"
                    QMessageBox.information(self, "成功", message)
                    self.accept()
                else:
                    type_str = "永久会员" if self.is_permanent else "年费会员"
                    logging.warning(f"激活码类型不匹配，需要{type_str}激活码")
                    QMessageBox.warning(self, "错误", f"请使用{type_str}激活码！")
            else:
                logging.warning("激活码无效")
                QMessageBox.warning(self, "错误", "激活码无效！")
                
        except Exception as e:
            logging.error(f"注册处理失败: {str(e)}")
            logging.error(traceback.format_exc())
            QMessageBox.critical(self, "错误", f"注册失败：{str(e)}") 