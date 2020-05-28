import {  } from "ts-jest";

import * as React from "react";
import { configure,mount,ReactWrapper } from "enzyme";
import * as adapter from "enzyme-adapter-react-16";

import ReactGroupSamples from "../components/ReactGroupSamples";
import {IReactGroupSamplesProps} from "../components/IReactGroupSamplesProps";
import { IReactGroupSampleState } from "../components/IReactGroupSampleState";

configure({adapter:new adapter()});
describe("React Group Samples Test Suite", () => {
 let reactWrapper:ReactWrapper<IReactGroupSamplesProps,IReactGroupSampleState>;
 
 beforeAll(()=>{

 });
    test("it should do something", () => {
    expect(1+1).toBe(2);
    });
});