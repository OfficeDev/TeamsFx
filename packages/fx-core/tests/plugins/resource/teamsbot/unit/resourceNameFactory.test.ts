// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import 'mocha';
import * as chai from 'chai';
import * as sinon from 'sinon';

import { ResourceNameFactory } from '../../src/utils/resourceNameFactory';
import * as utils from '../../src/utils/common';

describe('Resource Name Factory', () => {
    describe('createCommonName', () => {
        it('Happy Path', () => {
            // Arrange
            const appName = 'demo0329';
            const limit: number = 10;

            sinon.stub(utils, 'genUUID').returns('abcdefg');

            // Act

            const name = ResourceNameFactory.createCommonName(appName, limit);

            console.log(name);
            // Assert
            const expectName = 'tbpabcdefg';
            chai.assert.lengthOf(name, limit);
            chai.assert.isTrue(name === expectName);
        });
    });
});