// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Context, Middleware } from '@microsoft/microsoft-graph-client';

export default class ClientLoggingMiddleware implements Middleware {
  private nextMiddleware?: Middleware;

  // Work in progress - not currently very useful
  public async execute(context: Context): Promise<void> {
    console.log(context.request);
    if (this.nextMiddleware) {
      await this.nextMiddleware.execute(context);
    }
  }

  public setNext(middleware: Middleware): void {
    this.nextMiddleware = middleware;
  }
}
