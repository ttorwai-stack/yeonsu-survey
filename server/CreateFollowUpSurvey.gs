/**
 * [5급 승진예정자과정] 현업활용도 설문 만들기 (한 번만 실행)
 * Prompted by kidsreturn_79 · © 2026.10
 *
 * 쓰는 법: Apps Script 편집기 왼쪽 파일 목록 + → 스크립트 → 이름 "CreateFollowUpSurvey" → 이 내용을 붙여넣고 저장.
 *          위쪽 함수 선택에서 createFollowUpSurvey 를 고르고 [실행]. 만들어지는 즉시 "진행 중" 상태로 열린다.
 * 같은 이름의 설문이 이미 있으면 만들지 않는다(중복 방지).
 */
function createFollowUpSurvey() {
  var cfg = {
    v: 2,
    title: '[5급 승진예정자과정] 현업활용도 설문조사',
    course: '5급 승진예정자과정',
    cohort: '',
    period: '',
    intro: '안녕하세요? 연수 참여하시느라 애쓰셨습니다.\n\n' +
      '연수생의 의견을 수렴하여 보다 나은 연수를 기획하고자 연수 사후 현업활용도 설문 조사를 실시하고자 합니다. ' +
      '귀하의 응답은 추후 연수를 기획하고 운영하는 데에 소중한 밑거름이 될 것입니다. ' +
      '진솔하고 성실하게 응답해주시면 감사하겠습니다. (30초 정도 소요)\n\n' +
      '※ 응답하신 내용은 무기명으로 전산 처리되고, 통계법 제 33조(비밀의 보호)에 따라 비밀이 보장됩니다.',
    doneMsg: '소중한 의견 감사합니다.',
    sections: [
      {
        id: 'use', title: '현업활용도', desc: '교육 과정 각 평가 지표에 대해 표시해 주세요.',
        blocks: [{
          type: 'scale', id: 'use', title: '', area: '현업활용도', prefix: '현업활용도',
          labels: ['전혀 아니다', '그렇지 않다', '보통이다', '그렇다', '매우 그렇다'], posName: '긍정 비율',
          items: [
            { short: '업무 문제 해결', text: '본 교육과정(5급 승진예정자과정)에서 배운 지식과 기술은 앞으로 나의 업무와 관련된 문제를 해결하는 데 도움이 될 것으로 생각한다.' },
            { short: '실무 활용', text: '나는 본 교육과정에서 배운 내용을 실무에 효과적으로 활용하고 있다.' },
            { short: '업무성과 향상', text: '본 교육과정은 나의 업무성과 향상에 도움이 되었다.' }
          ]
        }]
      },
      {
        id: 'open', title: '의견', desc: '구체적으로 적어 주시면 다음 연수에 반영하겠습니다.',
        blocks: [
          { type: 'text', id: 'helped', short: '도움이 된 점', text: '교육과정에서 학습한 내용을 업무에 적용하는 데 있어 도움이 되었던 점이 있었다면 무엇인지 기술해 주세요.', required: false },
          { type: 'text', id: 'improve', short: '개선할 점', text: '업무 적용도를 향상시키기 위해서 개선해야 할 점이 있다면 무엇인지 기술해 주십시오.', required: false }
        ]
      }
    ]
  };

  var dup = listSurveys_(true).filter(function (m) { return m.title === cfg.title; })[0];
  if (dup) { Logger.log('이미 있어요: ' + dup.id + ' (' + dup.status + ')'); return dup.id; }
  var r = createSurveyFromConfig_(cfg);
  Logger.log('만들었어요. 설문ID: ' + r.id + ' / 진행 중');
  Logger.log('관리자 페이지의 설문 목록에서 [링크 복사]로 응답 링크를 받으세요.');
  return r.id;
}
