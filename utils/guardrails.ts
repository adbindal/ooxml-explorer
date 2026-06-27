/**
 * Guardrail keywords for detecting prompt injection or instruction leakage attempts.
 */
export const INJECTION_KEYWORDS = [
  'ignore previous instructions',
  'reveal system prompt',
  'system instructions',
  'how are you programmed',
  'spell check your instructions',
  'reveal your instructions',
  'what is your prompt',
  'what is your system prompt'
];

/**
 * Validates the input prompt to prevent prompt injection.
 * Throws a GUARDRAIL_VIOLATION error if an injection signature is detected.
 */
export const validateInput = (input: string): void => {
  const lowerInput = input.toLowerCase();
  for (const keyword of INJECTION_KEYWORDS) {
    if (lowerInput.includes(keyword)) {
      throw new Error("GUARDRAIL_VIOLATION: Input query contains blocked terms or injection patterns.");
    }
  }
};

/**
 * Validates the model output to prevent leakage of internal instructions or prompts.
 * Throws a GUARDRAIL_VIOLATION error if a leakage signature is detected.
 */
export const validateOutput = (output: string): void => {
  const lowerOutput = output.toLowerCase();
  // If the model starts repeating or referencing system prompt indicators, block it.
  if (
    lowerOutput.includes('system instructions') || 
    lowerOutput.includes('scope guard') || 
    lowerOutput.includes('source protection')
  ) {
    throw new Error("GUARDRAIL_VIOLATION: Output blocked by security shield due to potential instruction leakage.");
  }
};
