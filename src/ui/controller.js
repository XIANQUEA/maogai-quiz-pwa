export const selected=()=>[...document.querySelectorAll('input[name="answer"]:checked')].map(input=>input.value);
export const shouldSubmitOnSelection=type=>type==="judgment"||type==="single";
