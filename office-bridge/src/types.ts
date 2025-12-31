// WebSocket message types

export interface RegisterMessage {
  type: 'register';
  payload: {
    token: string;
    filename: string;
    path: string;
  };
}

export interface RegisteredMessage {
  type: 'registered';
  payload: {
    id: string;
  };
}

export interface ExecuteMessage {
  type: 'execute';
  id: string;
  payload: {
    code: string;
    timeout: number;
  };
}

export interface ResultMessage {
  type: 'result';
  id: string;
  payload: {
    success: boolean;
    result?: unknown;
    error?: {
      message: string;
      code?: string;
      stack?: string;
    };
  };
}

export interface ConsoleMessage {
  type: 'console';
  payload: {
    level: 'log' | 'warn' | 'error' | 'info';
    message: string;
    timestamp: string;
  };
}

export type AddInMessage = RegisterMessage | ResultMessage | ConsoleMessage;
export type BridgeMessage = RegisteredMessage | ExecuteMessage;

// Document registry types

export interface ConnectedDocument {
  id: string;
  filename: string;
  path: string;
  connectedAt: Date;
  lastActivity: Date;
  status: 'connected' | 'disconnected';
  disconnectedAt?: Date;
}

export interface DocumentInfo {
  id: string;
  filename: string;
  path: string;
  connectedAt: string;
  lastActivity: string;
  status: 'connected' | 'disconnected';
}

// Server info response
export interface ServerInfo {
  port: number;
  token: string;
  documents: number;
}
