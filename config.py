# Excel字段映射配置
EXCEL_MAPPING = {
    'employee_id': '工号',
    'name': '姓名',
    'work_type': '工作类型',
    'project_name': '项目名称/入池机构-项目名称',
    'project_stage': '项目阶段',
    'last_week_work': '上周三至本周二工作内容',
    'next_week_plan': '本周三至下周二工作计划',
    'issues': '问题反馈',
    'resume_count': '通过简历数量',
    'interview_count': '面试人员数量',
    'interview_pass_count': '面试通过人员数量'
}

# PDF格式配置
PDF_CONFIG = {
    'title': {
        'font': 'SourceHanSans',
        'size': 14,
        'alignment': 'center'
    },
    'subtitle': {
        'font': 'SourceHanSerif',
        'size': 10,
        'alignment': 'left'
    },
    'content': {
        'font': 'SourceHanSerif',
        'size': 9,
        'alignment': 'left'
    },
    'margins': {
        'left': 72,
        'right': 72,
        'top': 72,
        'bottom': 72
    }
}

# 项目阶段映射
PROJECT_STAGES = {
    '调研阶段': '（调研阶段）',
    '已立项进行中': '（已立项进行中）',
    '开发迭代中': '（开发迭代中）',
    '测试阶段': '（测试阶段）',
    '上线准备': '（上线准备）',
    '已上线': '（已上线）'
}

# 工作类型映射
WORK_TYPES = {
    '入池': '入池工作',
    '入项': '综合业务组项目'
} 