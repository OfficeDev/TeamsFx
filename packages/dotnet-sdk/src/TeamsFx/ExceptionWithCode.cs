﻿// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;

namespace Microsoft.TeamsFx
{
    /// <summary>
    /// Exception code to trace the exception types.
    /// </summary>
    public enum ExceptionCode
    {
        /// <summary>
        /// Invalid parameter error.
        /// </summary>
        InvalidParameter,
        /// <summary>
        /// Invalid configuration error.
        /// </summary>
        InvalidConfiguration,
        /// <summary>
        /// Internal error.
        /// </summary>
        InternalError,
        /// <summary>
        /// Channel is not supported error.
        /// </summary>
        ChannelNotSupported,
        /// <summary>
        /// Runtime is not supported error.
        /// </summary>
        RuntimeNotSupported,
        /// <summary>
        /// User failed to finish the AAD consent flow failed.
        /// </summary>
        ConsentFailed,
        /// <summary>
        /// The user or administrator has not consented to use the application error.
        /// </summary>
        UiRequiredError,
        /// <summary>
        /// Token is not within its valid time range error.
        /// </summary>
        TokenExpiredError,
        /// <summary>
        /// Call service (AAD or simple authentication server) failed.
        /// </summary>
        ServiceError,
        /// <summary>
        /// Operation failed.
        /// </summary>
        FailedOperation,
        /// <summary>
        /// General JSException that are not generated by TeamsFx SDK. 
        /// </summary>
        JSRuntimeError,
    }

    /// <summary>
    /// Exception class with code and message thrown by the SDK.
    /// </summary>
    public class ExceptionWithCode : Exception
    {
        /// <summary>
        /// Exception Code.
        /// </summary>
        public readonly ExceptionCode Code;

        internal ExceptionWithCode(string message, ExceptionCode exceptionCode) : base(message)
        {
            Code = exceptionCode;
        }
    }
}
