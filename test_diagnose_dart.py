"""diagnose_dart.py 판정 로직 회귀 테스트.

진단 도구가 세 경우를 각각 다르게 판정해야 한다.
잘못 판정하면 엉뚱한 해결책(국내 서버 임차 등)으로 이어지므로 여기가 핵심이다.
"""

import sys

import diagnose_dart as D

PASS, FAIL = [], []


def check(name, cond, detail=''):
    (PASS if cond else FAIL).append(name)
    print(f"  [{'OK ' if cond else 'FAIL'}] {name} {detail}")


print("=== 1. DART가 정상 응답한 경우 → '정상' ===")
# corpCode.xml 은 키가 틀리면 XML 로 status 를 돌려준다
v, d = D.classify({'body': '<?xml version="1.0"?><result><status>010</status>'
                           '<message>등록되지 않은 키입니다.</message></result>'})
check('XML status 응답 → 정상', v == '정상', f'→ {v}')

v, d = D.classify({'body': '{"status":"013","message":"조회된 데이터가 없습니다."}'})
check('JSON status 응답 → 정상', v == '정상', f'→ {v}')

v, d = D.classify({'body': 'PK\x03\x04\x14\x00'})   # 정상 corpCode 는 ZIP 으로 온다
check('ZIP(정상 데이터) 응답 → 정상', v == '정상', f'→ {v}')

print("\n=== 2. 연결이 안 된 경우 → '연결실패' ===")
v, d = D.classify({'error_kind': 'timeout', 'error': 'timeout'})
check('timeout → 연결실패', v == '연결실패', f'→ {v}')
check('제한시간을 안내에 포함', str(D.TIMEOUT) in d, f'→ {d[:40]}')

v, d = D.classify({'error_kind': 'unreachable', 'error': 'Connection refused'})
check('연결 거부 → 연결실패', v == '연결실패', f'→ {v}')

v, d = D.classify({'error_kind': 'other', 'error': 'SSL error'})
check('기타 오류 → 연결실패', v == '연결실패', f'→ {v}')
check('오류 내용을 그대로 전달', 'SSL error' in d, f'→ {d}')

print("\n=== 3. 차단 페이지가 온 경우 → '차단의심' ===")
v, d = D.classify({'body': '<html><head><title>Access Denied</title></head>'})
check('HTML 응답 → 차단의심', v == '차단의심', f'→ {v}')

v, d = D.classify({'body': '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">'})
check('DOCTYPE 대문자도 인식', v == '차단의심', f'→ {v}')

v, d = D.classify({'body': '\n\n   <html>차단되었습니다</html>'})
check('앞 공백 있어도 인식', v == '차단의심', f'→ {v}')

print("\n=== 4. 알 수 없는 응답 → '판단불가' ===")
v, d = D.classify({'body': 'something unexpected'})
check('정체불명 응답 → 판단불가', v == '판단불가', f'→ {v}')
v, d = D.classify({'body': ''})
check('빈 응답 → 판단불가', v == '판단불가', f'→ {v}')

print("\n=== 5. 네 판정이 서로 구분되는가 ===")
verdicts = {
    D.classify({'body': '<status>000</status>'})[0],
    D.classify({'error_kind': 'timeout'})[0],
    D.classify({'body': '<html>'})[0],
    D.classify({'body': 'zzz'})[0],
}
check('판정 4종이 모두 다름', len(verdicts) == 4, f'→ {sorted(verdicts)}')

print("\n=== 6. 실제 실행 (이 환경은 DART 차단됨 → 연결실패가 나와야 정상) ===")
probe = D.probe_dart(timeout=8)
v, d = D.classify(probe)
check('막힌 환경에서 연결실패로 판정', v == '연결실패', f'→ {v}: {d[:50]}')
check('소요 시간 기록', probe['elapsed'] is not None, f"→ {probe['elapsed']}초")

print("\n=== 7. 복사용 보고서 ===")
report = D.build_report({'ip': '1.2.3.4', 'country': 'KR', 'org': 'SK Broadband'},
                        None, '211.253.1.1', None,
                        {'url': D.DART_API, 'elapsed': 0.4, 'status_code': 200,
                         'content_type': 'application/xml', 'body': '<status>010</status>',
                         'error': None, 'error_kind': None},
                        '정상', 'DART가 정상 응답했습니다.')
for must in ['판정', '공인 IP', '국가', 'DNS', 'HTTP 코드', '응답 앞부분']:
    check(f"보고서에 '{must}' 포함", must in report)
check('국가 값이 들어감', 'KR' in report)

print("\n=== 8. IP 조회가 실패해도 진단이 계속되는가 ===")
r2 = D.build_report(None, '연결 안 됨', None, 'DNS 실패',
                    {'url': D.DART_API, 'elapsed': 15.0, 'status_code': None,
                     'body': None, 'error': 'timeout', 'error_kind': 'timeout'},
                    '연결실패', '응답 없음')
check('IP 실패해도 보고서 생성', '조회 실패' in r2 and '연결실패' in r2)

print("\n" + "=" * 56)
print(f"통과 {len(PASS)}건 / 실패 {len(FAIL)}건")
if FAIL:
    for f in FAIL:
        print(f"  - {f}")
    sys.exit(1)
print("전부 통과")
