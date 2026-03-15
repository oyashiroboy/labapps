require "test_helper"

class PptToolControllerTest < ActionDispatch::IntegrationTest
  test "should get index" do
    get ppt_tool_index_url
    assert_response :success
  end

  test "should get manual" do
    get ppt_tool_manual_url
    assert_response :success
  end
end
