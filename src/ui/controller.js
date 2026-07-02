export const selected=()=>[...document.querySelectorAll('input[name="answer"]:checked')].map(input=>input.value);
