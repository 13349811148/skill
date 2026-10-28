"""Persist the browser upload handoff; never claim that a local export is online."""
import argparse
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse
from zipfile import ZipFile


def sidecar(report):
    return Path(report).with_suffix('.wps-upload.json')


def save(path, data):
    temporary = path.with_name(path.name + '.tmp')
    temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    temporary.replace(path)


def digest(report):
    return hashlib.sha256(report.read_bytes()).hexdigest()


def prepare(report, platform, generated_after_ns=0):
    report = Path(report).resolve()
    if not report.is_file() or report.stat().st_mtime_ns < generated_after_ns:
        raise ValueError('商品日报没有在本次运行中成功导出，不能上传旧文件')
    with ZipFile(report) as archive:
        if 'xl/workbook.xml' not in archive.namelist() or archive.testzip():
            raise ValueError('商品日报不是完整可读取的 XLSX 文件')
    checksum = digest(report)
    path = sidecar(report)
    previous = json.loads(path.read_text(encoding='utf-8')) if path.exists() else {}
    if previous.get('sha256') == checksum and previous.get('status') == 'uploaded':
        return previous
    data = {'schema_version': 1, 'report': str(report), 'platform': platform,
            'sha256': checksum, 'status': 'pending', 'destination': 'WPS 我的云文档',
            'url': None, 'message': '本地已生成，等待代理执行 WPS 浏览器上传并验证',
            'updated_at': datetime.now(timezone.utc).isoformat()}
    data['previous_url']=previous.get('url') or previous.get('previous_url')
    save(path, data)
    return data


def record(report, status, url=None, message=''):
    report = Path(report).resolve()
    path = sidecar(report)
    data = json.loads(path.read_text(encoding='utf-8'))
    if data['sha256'] != digest(report):
        raise ValueError('本地报表已变更，请先重新 prepare，不能沿用旧上传状态')
    if status == 'uploaded':
        parsed = urlparse(url or '')
        if parsed.scheme != 'https' or not (parsed.hostname == 'kdocs.cn' or (parsed.hostname or '').endswith('.kdocs.cn')) or not parsed.path.startswith('/l/') or len(parsed.path) <= 3:
            raise ValueError('上传成功必须提供浏览器实际验证的 HTTPS WPS 文档链接')
    data.update(status=status, url=url if status == 'uploaded' else None,
                message=message, updated_at=datetime.now(timezone.utc).isoformat())
    save(path, data)
    return data


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=['prepare', 'record'])
    parser.add_argument('--report', required=True, type=Path)
    parser.add_argument('--platform', default='pdd', choices=['pdd', 'tmall'])
    parser.add_argument('--status', choices=['pending', 'login_required', 'failed', 'uploaded'])
    parser.add_argument('--url')
    parser.add_argument('--message', default='')
    args = parser.parse_args()
    if args.action == 'record' and not args.status:
        parser.error('record requires --status')
    data = prepare(args.report, args.platform) if args.action == 'prepare' else record(args.report, args.status, args.url, args.message)
    print(json.dumps(data, ensure_ascii=False))


if __name__ == '__main__':
    main()
