/**
 * PromptBuilder.js — دامج System + Task
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.PromptBuilder = {
  /**
   * @param {string} taskType - "parser" | "validator" | "reporter" | "risk"
   * @param {Object} ctx - context للمهمة
   * @return {Object} { system, user }
   */
  build: function(taskType, ctx) {
    var taskFn = GDS2.TaskPrompts[taskType];
    if (typeof taskFn !== 'function') {
      throw new Error('Unknown task type: ' + taskType);
    }
    return {
      system: GDS2.SystemPrompt.BASE,
      user: taskFn(ctx || {})
    };
  }
};
