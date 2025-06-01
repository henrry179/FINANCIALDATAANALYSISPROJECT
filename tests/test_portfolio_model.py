from src.models.portfolio_model import PortfolioModel

def test_portfolio_model_init():
    data = [1, 2, 3]
    model = PortfolioModel(data)
    assert model.data == data

def test_portfolio_model_optimize():
    data = [1, 2, 3]
    model = PortfolioModel(data)
    assert hasattr(model, 'optimize') 