import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import DelegatedGraphCommand from '../../../base/DelegatedGraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
}

class TodoListAddCommand extends DelegatedGraphCommand {
  public get name(): string {
    return commands.LIST_ADD;
  }

  public get description(): string {
    return 'Adds a new Microsoft To Do task list';
  }

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const data: any = {
      displayName: args.options.name
    };

    const requestOptions: any = {
      url: `${this.resource}/v1.0/me/todo/lists`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      data,
      responseType: 'json'
    };

    try {
      const res: any = await request.post(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TodoListAddCommand();