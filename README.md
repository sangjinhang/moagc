# COSMIC 需求文档生成器 - 云部署

## 快速部署（Docker）

```bash
# 1. 构建镜像
docker build -t cosmic-generator .

# 2. 运行容器
docker run -d -p 8080:8080 --name cosmic cosmic-generator

# 3. 访问
# http://你的公网IP:8080
```

## 阿里云/腾讯云 ECS 部署

### 方式一：Docker（推荐）

```bash
# SSH 登录服务器
ssh root@你的IP

# 安装 Docker
curl -fsSL https://get.docker.com | sh

# 上传代码（在本地执行）
scp -r deploy/ root@你的IP:/opt/cosmic/

# 在服务器上构建运行
cd /opt/cosmic
docker build -t cosmic .
docker run -d -p 80:8080 --restart=always --name cosmic cosmic
```

### 方式二：直接运行

```bash
# 在服务器上
cd /opt/cosmic
pip install -r requirements.txt
gunicorn --bind 0.0.0.0:80 --workers 2 --threads 4 --timeout 600 app:app
```

## 安全组/防火墙

确保开放端口：
- 阿里云：安全组 → 添加入方向规则 → TCP 80
- 腾讯云：安全组 → 添加规则 → TCP 80

## 文件结构

```
deploy/
├── app.py              # Flask 主应用
├── requirements.txt    # Python 依赖
├── Dockerfile          # Docker 构建文件
├── COSMIC模板.xlsx     # 模板文件
└── scripts/            # Pipeline 脚本
    ├── pipeline.py
    ├── convert_cosmic13.py
    ├── draw_sequence.py
    ├── generate_fsd_excel.py
    ├── convert_fsd_to_images.py
    └── rewrite_h5_to_prd.py
```
