// ─── app.js ───
// 네비게이션(hash 기반 화면 전환) + 전역 이벤트 바인딩

const VALID_SCREENS = ['home', 'excel', 'merge'];

function renderScreen(name) {
  $('.screen').removeClass('active');
  $('.tool-block').removeClass('active');
  $(`#screen-${name}`).addClass('active');
  $(`#nav-${name}`).addClass('active');
  $('.main').scrollTop(0);
}

export function showScreen(name) {
  if (!VALID_SCREENS.includes(name)) name = 'home';
  location.hash = name;
}

export function goBack() {
  history.back();
}

$(window).on('hashchange', () => {
  const name = location.hash.replace('#', '') || 'home';
  renderScreen(VALID_SCREENS.includes(name) ? name : 'home');
});

// 초기 화면 세팅
const initScreen = location.hash.replace('#', '') || 'home';
renderScreen(VALID_SCREENS.includes(initScreen) ? initScreen : 'home');
if (!location.hash) location.replace('#home');
