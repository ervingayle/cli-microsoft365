export const prompt = {
  /* c8 ignore next 10 */
  async forInput<T>(config: any, answers?: any): Promise<T> {
    const inquirer: any = require('inquirer');

    const response = await inquirer.prompt(config, answers) as T;

    return response;
  }
};